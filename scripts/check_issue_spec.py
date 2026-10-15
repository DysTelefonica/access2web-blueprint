#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — assets/scripts/check_issue_spec.py
"""PR body traceability gate (issue #855).

The branch name already carries its issue number (`check_branch_name.py`).
A reviewer who reads only the PR body still cannot tell which issue the PR
closes: branch names drift, copy/paste loses suffixes, and chain slices
republish under different PR numbers. The body has to speak for itself.

Contract:

- The body must contain ``Closes|Fixes|Resolves|Refs #N`` for at least one
  issue number.
- ``GET /repos/{owner}/{repo}/issues/{n}`` with ``GITHUB_TOKEN`` must return
  a real issue (not a pull request) carrying ``status:approved`` and at
  least one ``type:*`` label.
- An unreadable API response is fail-closed: the gate exits 1 rather than
  silently accepting the reference.

Two entry points:

- ``local --body-file <path> --repository <owner/repo>`` — direct body,
  no GitHub event payload needed.
- ``pr-event <event.json>`` — pull-request body and repository are read
  from a GitHub webhook payload, the shape produced by ``GITHUB_EVENT_PATH``.

Exit codes:

- 0  body references a valid approved+typed issue.
- 1  anything else: missing reference, reference to a PR, missing labels,
   API error, malformed event payload.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.request
from collections.abc import Mapping, Sequence
from typing import Any, Protocol

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------

APPROVAL_LABEL = "status:approved"
TYPE_PREFIX = "type:"

# Issue #855: the closing-keyword family used by GitHub, plus ``Refs`` for
# non-closing references in chain slices. The check is case-insensitive and
# parses prose verbatim — that is the contract that frees the branch name
# from carrying the issue number (the previous coupling forced manual
# republishes when a chain slice changed).
_REFERENCED_ISSUE_RE = re.compile(r"(?i)\b(?:closes|fixes|resolves|refs)\s+#(?P<number>\d+)\b")


# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------


class GitHubApiError(RuntimeError):
    """Raised when GitHub cannot provide authoritative issue data."""


class PullRequestClient(Protocol):
    """The read-only GitHub surface the PR traceability gate needs."""

    def issue(self, repository: str, number: int) -> Mapping[str, Any]: ...


def referenced_issue_numbers(body: str) -> tuple[int, ...]:
    """Return deduplicated, sorted issue numbers referenced in the PR body.

    Accepts the GitHub closing-keyword family (``Closes``, ``Fixes``,
    ``Resolves``) plus ``Refs`` for non-closing references. The check is
    case-insensitive and parses prose verbatim. Empty / None bodies yield
    an empty tuple — the gate fails on the no-reference rule, not here.
    """
    if not body:
        return ()
    return tuple(
        sorted({int(match.group("number")) for match in _REFERENCED_ISSUE_RE.finditer(body)})
    )


def _labels(issue: Mapping[str, Any]) -> set[str]:
    return {
        label["name"] if isinstance(label, Mapping) else str(label)
        for label in issue.get("labels", [])
    }


def validate_issue(issue: Mapping[str, Any], number: int) -> list[str]:
    """Return contract violations for one GitHub issue.

    The gate is intentionally short: a PR reference, a missing approval
    label, or a missing ``type:*`` label. Anything beyond that (the full
    six-section spec APAP_WEB enforces) is out of scope for issue #855 —
    ``cuanto más simple sin que deje de comprobar lo esencial mejor``.
    """
    if "pull_request" in issue:
        return [f"#{number}: reference resolves to a pull request, not an issue"]
    labels = _labels(issue)
    if APPROVAL_LABEL not in labels:
        return [f"#{number}: requires {APPROVAL_LABEL}"]
    if not any(label.startswith(TYPE_PREFIX) for label in labels):
        return [f"#{number}: requires a {TYPE_PREFIX}* label"]
    return []


def validate_body(
    body: str,
    client: PullRequestClient,
    repository: str,
) -> list[str]:
    """Validate every body-referenced issue against the GitHub issue surface."""
    numbers = referenced_issue_numbers(body)
    if not numbers:
        return [
            "PR body must reference at least one issue with 'Closes #N', 'Fixes #N', "
            "'Resolves #N' or 'Refs #N' (pointing to an existing issue, not a pull request)"
        ]
    violations: list[str] = []
    for number in numbers:
        try:
            issue = client.issue(repository, number)
        except GitHubApiError as exc:
            # Fail-closed: a missing or unreadable issue is a contract
            # violation to report, not a crash of the gate.
            violations.append(f"#{number}: cannot read the issue ({exc})")
            continue
        violations.extend(validate_issue(issue, number))
    return violations


def validate_pr_event(event: Mapping[str, Any], client: PullRequestClient) -> list[str]:
    """Validate a GitHub ``pull_request`` webhook payload.

    Reads the body and the repository from the event and forwards to
    ``validate_body``. A payload that lacks ``repository.full_name`` cannot
    be resolved and is rejected up front — anything that reaches the API
    client must already carry the target identity.
    """
    repository = (event.get("repository") or {}).get("full_name", "")
    if not repository or "/" not in repository:
        return ["event does not identify repository.full_name"]
    pull_request = event.get("pull_request") or {}
    body = pull_request.get("body") or ""
    return validate_body(body, client, repository)


class GitHubClient:
    """Minimal read-only GitHub REST client.

    In a public repo, the token only needs to read issues — the workflow
    file grants ``issues: read`` on the job that runs this gate. A
    fork-only ``GITHUB_TOKEN`` is enough.
    """

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
        # The base URL is fixed by GitHub Actions or the caller; arbitrary
        # user-supplied schemes never enter this boundary.
        request = urllib.request.Request(  # noqa: S310
            f"{self._api_url}/repos/{repository}/issues/{number}",
            headers=headers,
        )
        try:
            with urllib.request.urlopen(request, timeout=30) as response:  # noqa: S310
                return json.load(response)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
            raise GitHubApiError(f"GET /repos/{repository}/issues/{number} failed: {exc}") from exc


def _report(violations: Sequence[str]) -> int:
    """Print each violation to stderr and return the gate exit code."""
    for violation in violations:
        print(f"issue-spec: {violation}", file=sys.stderr)
    if violations:
        print(f"issue-spec: failed with {len(violations)} violation(s)", file=sys.stderr)
        return 1
    print("issue-spec: contract satisfied")
    return 0


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8 on every supported workstation locale.

    Python picks the output encoding from the platform locale, so the same
    gate emits different bytes on a Windows workstation (cp1252) and a Linux
    runner (utf-8) — and a non-encodable character crashes the write
    outright. A harness that claims determinism cannot let its own output
    depend on where it ran.
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


# --------------------------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------------------------


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    local = subparsers.add_parser(
        "local",
        help="validate a PR body read directly from a file",
    )
    local.add_argument(
        "--body-file",
        type=argparse.FileType("r", encoding="utf-8"),
        required=True,
        help="path to the PR body (markdown)",
    )
    local.add_argument(
        "--repository",
        required=True,
        help="GitHub repository in 'owner/name' form",
    )

    pr_event = subparsers.add_parser(
        "pr-event",
        help="validate issues linked by a GitHub PR webhook event payload",
    )
    pr_event.add_argument(
        "event",
        type=argparse.FileType("r", encoding="utf-8"),
        help="path to the GitHub PR event JSON (e.g. $GITHUB_EVENT_PATH)",
    )
    return parser


# Exposed for tests: parsing must be reachable without triggering the API.
parse_args = _build_parser().parse_args


def _make_client() -> PullRequestClient:
    return GitHubClient(
        token=os.environ.get("GITHUB_TOKEN", ""),
        api_url=os.environ.get("GITHUB_API_URL", "https://api.github.com"),
    )


def run_command(args: argparse.Namespace, client: PullRequestClient) -> int:
    """Dispatch one parsed invocation to the right validator.

    Splitting ``main`` from ``run_command`` keeps the network boundary at
    the door and lets unit tests inject a stub client in one line.
    """
    if args.command == "local":
        body = args.body_file.read()
        return _report(validate_body(body, client, args.repository))
    event = json.load(args.event)
    return _report(validate_pr_event(event, client))


def main(argv: Sequence[str] | None = None) -> int:
    _pin_output_encoding()
    args = _build_parser().parse_args(argv)
    return run_command(args, _make_client())


if __name__ == "__main__":
    raise SystemExit(main())
