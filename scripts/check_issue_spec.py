#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — assets/scripts/check_issue_spec.py
"""PR body traceability gate (issue #855).

The PR body must reference an issue that is approved and typed. Without this
gate the issue only appears in the branch name — and the branch is not
readable when the reviewer needs it.

Rules:

1. Body without ``Closes|Fixes|Resolves|Refs #N`` → exit 1.
2. Reference resolves to a pull request, not an issue → exit 1.
3. Issue lacks ``status:approved`` or any ``type:*`` label → exit 1.
4. Valid issue (approved + typed) → exit 0.
5. GitHub API error → exit 1 (fail-closed).

Modes:

- ``local --body-file <path> --repository <owner/repo>`` — direct body.
- ``pr-event <event.json>`` — GitHub PR webhook payload.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.request
from collections.abc import Callable, Mapping
from typing import Any

APPROVAL_LABEL = "status:approved"
TYPE_PREFIX = "type:"
_REFERENCED_ISSUE_RE = re.compile(r"(?i)\b(?:closes|fixes|resolves|refs)\s+#(\d+)\b")


class GitHubApiError(RuntimeError):
    """Raised when GitHub cannot provide authoritative issue data."""


# --------------------------------------------------------------------------------------------
# One function per rule.
# --------------------------------------------------------------------------------------------


def referenced_issue_numbers(body: str) -> tuple[int, ...]:
    """Return deduplicated, sorted issue numbers from the body."""
    if not body:
        return ()
    return tuple(sorted({int(m.group(1)) for m in _REFERENCED_ISSUE_RE.finditer(body)}))


def rule_no_reference(body: str) -> str | None:
    """Rule 1: body must reference an issue with one of the four keywords."""
    if referenced_issue_numbers(body):
        return None
    return (
        "PR body must reference an issue with 'Closes #N', 'Fixes #N', 'Resolves #N' or 'Refs #N'"
    )


def rule_reference_is_pr(issue: Mapping[str, Any], number: int) -> str | None:
    """Rule 2: ``GET /issues/{n}`` returns ``pull_request`` for PRs, not issues."""
    if "pull_request" not in issue:
        return None
    return f"#{number}: reference resolves to a pull request, not an issue"


def rule_missing_labels(issue: Mapping[str, Any], number: int) -> list[str]:
    """Rule 3: issue must carry ``status:approved`` AND a ``type:*`` label."""
    names = {
        label["name"] if isinstance(label, Mapping) else str(label)
        for label in issue.get("labels", [])
    }
    violations: list[str] = []
    if APPROVAL_LABEL not in names:
        violations.append(f"#{number}: requires {APPROVAL_LABEL}")
    if not any(name.startswith(TYPE_PREFIX) for name in names):
        violations.append(f"#{number}: requires a {TYPE_PREFIX}* label")
    return violations


def rule_api_error(number: int, exc: GitHubApiError) -> str:
    """Rule 5: API error fails closed — surface it as a contract violation."""
    return f"#{number}: cannot read the issue ({exc})"


# --------------------------------------------------------------------------------------------
# Orchestration. Rule 4 is the absence of any of the above.
# --------------------------------------------------------------------------------------------


def validate(
    body: str,
    fetch_issue: Callable[[int], Mapping[str, Any]],
) -> list[str]:
    """Run every rule against the body and its referenced issues."""
    numbers = referenced_issue_numbers(body)
    if not numbers:
        return [rule_no_reference(body) or ""]
    violations: list[str] = []
    for number in numbers:
        try:
            issue = fetch_issue(number)
        except GitHubApiError as exc:
            violations.append(rule_api_error(number, exc))
            continue
        if msg := rule_reference_is_pr(issue, number):
            violations.append(msg)
        else:
            violations.extend(rule_missing_labels(issue, number))
    return [v for v in violations if v]


# --------------------------------------------------------------------------------------------
# GitHub client — read-only; ``issues: read`` is enough.
# --------------------------------------------------------------------------------------------


class GitHubClient:
    def __init__(self, token: str = "", api_url: str = "https://api.github.com") -> None:
        self._token = token
        self._api_url = api_url.rstrip("/")

    def issue(self, repository: str, number: int) -> Mapping[str, Any]:
        headers = {
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "access2web-blueprint-issue-spec",
        }
        if self._token:
            headers["Authorization"] = f"Bearer {self._token}"
        request = urllib.request.Request(  # noqa: S310
            f"{self._api_url}/repos/{repository}/issues/{number}", headers=headers
        )
        try:
            with urllib.request.urlopen(request, timeout=30) as response:  # noqa: S310
                return json.load(response)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
            raise GitHubApiError(f"GET /repos/{repository}/issues/{number} failed: {exc}") from exc


# --------------------------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------------------------


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    local = subparsers.add_parser("local", help="validate a body read from --body-file")
    local.add_argument(
        "--body-file",
        type=argparse.FileType("r", encoding="utf-8"),
        required=True,
    )
    local.add_argument("--repository", required=True)
    event = subparsers.add_parser("pr-event", help="validate a GitHub PR webhook event payload")
    event.add_argument("event", type=argparse.FileType("r", encoding="utf-8"))
    return parser


def _make_fetcher(repository: str) -> Callable[[int], Mapping[str, Any]]:
    """Build the issue fetcher from ``$GITHUB_TOKEN`` and ``$GITHUB_API_URL``.

    Tests monkeypatch this function to inject a stub without monkeypatching
    ``sys.modules``.
    """
    client = GitHubClient(
        token=os.environ.get("GITHUB_TOKEN", ""),
        api_url=os.environ.get("GITHUB_API_URL", "https://api.github.com"),
    )
    return lambda n: client.issue(repository, n)


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    if args.command == "local":
        body = args.body_file.read()
        repository = args.repository
    else:
        payload = json.load(args.event)
        repository = (payload.get("repository") or {}).get("full_name", "")
        if "/" not in repository:
            print("issue-spec: event does not identify repository.full_name", file=sys.stderr)
            return 1
        body = (payload.get("pull_request") or {}).get("body") or ""
    violations = validate(body, _make_fetcher(repository))
    for violation in violations:
        print(f"issue-spec: {violation}", file=sys.stderr)
    if violations:
        print(f"issue-spec: failed with {len(violations)} violation(s)", file=sys.stderr)
        return 1
    print("issue-spec: contract satisfied")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
