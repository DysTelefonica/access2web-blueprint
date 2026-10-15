"""Tests for ``scripts/check_issue_spec.py`` (issue #855).

Five rules, one test function each. Rules 1 and 3 are parametrized over the
multiple bodies / label combinations that share the same acceptance
criterion; the others run a single case each.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pytest

SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import check_issue_spec  # noqa: E402

REPO = "DysTelefonica/access2web-blueprint"


class StubClient:
    """In-memory fetcher for ``validate``: callable, raises ``GitHubApiError``."""

    def __init__(self, issues: dict[int, dict[str, Any]] | None = None) -> None:
        self.issues_by_number = dict(issues or {})

    def __call__(self, number: int) -> dict[str, Any]:
        if number not in self.issues_by_number:
            raise check_issue_spec.GitHubApiError(f"GET /repos/x/issues/{number}: 404")
        return self.issues_by_number[number]


def _issue(labels: list[str] | None = None) -> dict[str, Any]:
    return {"number": 42, "labels": [{"name": label} for label in (labels or [])]}


# Rule 1 — body without ``Closes/Fixes/Resolves/Refs #N`` → exit 1.
@pytest.mark.parametrize(
    "body",
    ["", "PR body without any reference.", "See #42 for prior art."],
)
def test_rule_1_no_reference(body: str) -> None:
    """Bare mentions and empty bodies count the same: no keyword, no contract."""
    fetcher = StubClient()
    assert check_issue_spec.validate(body, fetcher) == [check_issue_spec.rule_no_reference(body)]


# Rule 2 — reference resolves to a pull request → exit 1.
def test_rule_2_reference_is_pr() -> None:
    """A payload with ``pull_request`` set is a PR, not an issue."""
    issue = _issue(["status:approved", "type:feature"])
    issue["pull_request"] = {"url": "..."}
    fetcher = StubClient({42: issue})
    violations = check_issue_spec.validate("Closes #42", fetcher)
    assert any("pull request" in v for v in violations)


# Rule 3 — issue lacks ``status:approved`` or ``type:*`` → exit 1.
@pytest.mark.parametrize(
    "labels",
    [
        ["type:feature"],  # no approval
        ["status:approved"],  # no type
        [],  # neither
    ],
)
def test_rule_3_missing_labels(labels: list[str]) -> None:
    """Any of the three label configurations fails the gate."""
    fetcher = StubClient({42: _issue(labels)})
    violations = check_issue_spec.validate("Closes #42", fetcher)
    assert violations


# Rule 4 — valid issue → exit 0.
def test_rule_4_valid_issue_passes() -> None:
    """Approved + typed is the contract; nothing else needs to hold here."""
    fetcher = StubClient({42: _issue(["status:approved", "type:feature"])})
    assert check_issue_spec.validate("Closes #42", fetcher) == []


# Rule 5 — GitHub API error → exit 1 (fail-closed).
def test_rule_5_api_error_fails_closed() -> None:
    """A missing issue is a contract violation, not a silent pass."""
    fetcher = StubClient({})  # 404 on any number
    violations = check_issue_spec.validate("Closes #99", fetcher)
    assert any("cannot read" in v for v in violations)


# CLI smoke — the parser, the file argument and the exit code are wired.
def test_cli_local_mode(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    body_file = tmp_path / "body.md"
    body_file.write_text("Closes #42", encoding="utf-8")
    fetcher = StubClient({42: _issue(["status:approved", "type:feature"])})
    monkeypatch.setattr(check_issue_spec, "_make_fetcher", lambda repository: fetcher)
    rc = check_issue_spec.main(["local", "--body-file", str(body_file), "--repository", REPO])
    assert rc == 0
