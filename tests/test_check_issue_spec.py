# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — tests/test_check_issue_spec.py
"""Tests for `scripts/check_issue_spec.py` (issue #855).

Issue #855 closes the traceability gap left by `check_branch_name.py`: the
PR body must reference an issue that is actually approved and typed, not just
a branch that happens to carry a number. The acceptance contract is small
and binary — five rules, one exit code each — so the test suite mirrors it:
for each rule a positive case (the gate exits 0) and a negative case (the gate
exits 1 with the rule named in stderr). If any negative case passes, the gate
is silently broken and the next PR slips through.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

# Make `scripts/` importable without sys.path games: pytest adds the repo
# root, the script is at `scripts/check_issue_spec.py` from there.
SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import check_issue_spec  # noqa: E402  — sys.path tweak above

REPO = "DysTelefonica/access2web-blueprint"


# --------------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------------


class StubClient:
    """Minimal read-only GitHub stub for the PR traceability tests."""

    def __init__(self, issues: dict[int, dict[str, Any]] | None = None) -> None:
        self.issues_by_number = dict(issues or {})
        self.requested: list[tuple[str, int]] = []

    def issue(self, repository: str, number: int) -> dict[str, Any]:
        self.requested.append((repository, number))
        if number not in self.issues_by_number:
            raise check_issue_spec.GitHubApiError(f"GET /repos/{repository}/issues/{number}: 404")
        return self.issues_by_number[number]


def _issue(labels: list[str] | None = None) -> dict[str, Any]:
    """Build an issue with the labels the gate cares about."""
    return {
        "number": 42,
        "body": "x",
        "labels": [{"name": label} for label in (labels or [])],
    }


def _pr_event(body: str = "Closes #42", number: int = 7) -> dict[str, Any]:
    """Build a GitHub PR webhook payload with the body the caller cares about."""
    return {
        "repository": {"full_name": REPO},
        "pull_request": {
            "number": number,
            "user": {"login": "contributor"},
            "head": {"ref": "feat/42-slug"},
            "body": body,
        },
    }


# --------------------------------------------------------------------------------------------
# Acceptance criterion 1: body without a reference → exit 1
# --------------------------------------------------------------------------------------------


def test_body_without_reference_is_rejected() -> None:
    """No `Closes/Fixes/Resolves/Refs #N` in the body → the gate fails closed."""
    client = StubClient({42: _issue(["status:approved", "type:feature"])})

    violations = check_issue_spec.validate_body("PR body without any reference.", client, REPO)

    assert violations
    assert any("Closes" in v and "Fixes" in v for v in violations)
    assert client.requested == []


def test_main_local_mode_rejects_body_without_reference(tmp_path: Path) -> None:
    body_file = tmp_path / "body.md"
    body_file.write_text("PR body without any reference.", encoding="utf-8")
    client = StubClient({42: _issue(["status:approved", "type:feature"])})

    rc = check_issue_spec.run_command(
        check_issue_spec.parse_args(["local", "--body-file", str(body_file), "--repository", REPO]),
        client,
    )

    assert rc == 1


# --------------------------------------------------------------------------------------------
# Acceptance criterion 2: reference resolves to a PR → exit 1
# --------------------------------------------------------------------------------------------


def test_reference_to_a_pull_request_is_rejected() -> None:
    """`GET /issues/{n}` returns a payload with a `pull_request` key for PRs.

    The gate treats that as a contract violation: traceability requires an
    issue, not another PR.
    """
    pr_payload = _issue(labels=["status:approved"])
    pr_payload["pull_request"] = {"url": "https://api.github.com/.../pulls/42"}
    client = StubClient({42: pr_payload})

    violations = check_issue_spec.validate_body("Closes #42", client, REPO)

    assert any("pull request" in v for v in violations)


def test_main_pr_event_mode_rejects_reference_to_pull_request(tmp_path: Path) -> None:
    event_file = tmp_path / "event.json"
    event_file.write_text(json.dumps(_pr_event("Closes #42")), encoding="utf-8")
    pr_payload = _issue(labels=["status:approved"])
    pr_payload["pull_request"] = {"url": "..."}
    client = StubClient({42: pr_payload})

    rc = check_issue_spec.run_command(
        check_issue_spec.parse_args(["pr-event", str(event_file)]),
        client,
    )

    assert rc == 1


# --------------------------------------------------------------------------------------------
# Acceptance criterion 3: issue missing `status:approved` or `type:*` → exit 1
# --------------------------------------------------------------------------------------------


def test_issue_without_status_approved_is_rejected() -> None:
    """The gate is fail-closed on the approval label: a feature label alone is not enough."""
    client = StubClient({42: _issue(["type:feature"])})

    violations = check_issue_spec.validate_body("Closes #42", client, REPO)

    assert any("status:approved" in v for v in violations)


def test_issue_without_type_label_is_rejected() -> None:
    """A bare approval does not imply a type — both are required."""
    client = StubClient({42: _issue(["status:approved"])})

    violations = check_issue_spec.validate_body("Closes #42", client, REPO)

    assert any("type:" in v for v in violations)


def test_issue_with_no_labels_is_rejected() -> None:
    """No labels at all must fail on the first requirement the gate hits."""
    client = StubClient({42: _issue([])})

    violations = check_issue_spec.validate_body("Closes #42", client, REPO)

    assert any("status:approved" in v for v in violations)


# --------------------------------------------------------------------------------------------
# Acceptance criterion 4: valid issue → exit 0
# --------------------------------------------------------------------------------------------


def test_valid_issue_passes() -> None:
    """`status:approved` + `type:*` is the full contract."""
    client = StubClient({42: _issue(["status:approved", "type:feature"])})

    violations = check_issue_spec.validate_body("Closes #42", client, REPO)

    assert violations == []
    assert client.requested == [(REPO, 42)]


def test_case_insensitive_keywords_are_accepted() -> None:
    """`closes` lowercase and `Fixes` capitalized both count — the regex is case-insensitive."""
    client = StubClient({42: _issue(["status:approved", "type:feature"])})

    assert check_issue_spec.validate_body("closes #42", client, REPO) == []
    assert check_issue_spec.validate_body("Fixes #42", client, REPO) == []
    assert check_issue_spec.validate_body("Resolves #42", client, REPO) == []
    assert check_issue_spec.validate_body("Refs #42", client, REPO) == []


def test_multiple_references_are_each_validated() -> None:
    """Multiple `Closes #N` references each go through the gate."""
    client = StubClient(
        {
            1: _issue(["status:approved", "type:feature"]),
            2: _issue(["status:approved", "type:fix"]),
        }
    )

    violations = check_issue_spec.validate_body("Closes #1\nRefs #2", client, REPO)

    assert violations == []
    assert client.requested == [(REPO, 1), (REPO, 2)]


def test_main_pr_event_mode_passes_for_valid_issue(tmp_path: Path) -> None:
    event_file = tmp_path / "event.json"
    event_file.write_text(json.dumps(_pr_event("Closes #42")), encoding="utf-8")
    client = StubClient({42: _issue(["status:approved", "type:feature"])})

    rc = check_issue_spec.run_command(
        check_issue_spec.parse_args(["pr-event", str(event_file)]),
        client,
    )

    assert rc == 0


# --------------------------------------------------------------------------------------------
# Acceptance criterion 5: API error → exit 1 (fail-closed)
# --------------------------------------------------------------------------------------------


def test_api_error_fails_closed() -> None:
    """An exception in `client.issue` is reported, not swallowed.

    The default client raises `GitHubApiError` on transport / 4xx / 5xx; the
    gate treats any unreadable issue as a contract violation so a flaky
    GitHub never lets a missing reference pass.
    """
    client = StubClient({})  # No mapping → client.issue raises GitHubApiError

    violations = check_issue_spec.validate_body("Closes #99", client, REPO)

    assert any("cannot read" in v for v in violations)


# --------------------------------------------------------------------------------------------
# Reference parsing
# --------------------------------------------------------------------------------------------


def test_referenced_issue_numbers_dedupes_and_sorts() -> None:
    """Repeated mentions of the same number collapse; order is ascending."""
    body = "Closes #3, refs #1, fixes #3, Resolves #2"

    assert check_issue_spec.referenced_issue_numbers(body) == (1, 2, 3)


def test_referenced_issue_numbers_ignores_bare_mentions() -> None:
    """Bare `#N` (no closing keyword) is prose noise, not traceability."""
    body = "See #1421 for prior art; nothing to do with the body keyword here."

    assert check_issue_spec.referenced_issue_numbers(body) == ()


def test_referenced_issue_numbers_handles_empty_body() -> None:
    assert check_issue_spec.referenced_issue_numbers("") == ()


# --------------------------------------------------------------------------------------------
# Event payload sanity
# --------------------------------------------------------------------------------------------


def test_pr_event_without_repository_fails() -> None:
    """An event payload that lacks `repository.full_name` cannot be validated."""
    event = _pr_event()
    event["repository"] = {}
    client = StubClient({42: _issue(["status:approved", "type:feature"])})

    violations = check_issue_spec.validate_pr_event(event, client)

    assert any("repository.full_name" in v for v in violations)
