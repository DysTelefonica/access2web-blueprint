# HARNESS-PROVENANCE: access2web-blueprint issue #854 — test_pr_size_wiring.py
"""Wiring pin for `scripts/check_pr_size.py` after the #854 refactor.

The PR size gate counts additions + deletions and fails above 400 (gentle-ai
reference, `.github/scripts/check-pr-size.cjs:106-110`). Override is the
`size:exception` label only — no body reason line is required.

Acceptance rules under test (one test per rule):

    1. ≤400 passes (exit 0).
    2. >400 without label fails (exit 1) with the total in the message.
    3. >400 with `size:exception` label passes.
    4. `local` gives the same total as the API for the same range.
    5. Binary output, no warning.
"""

from __future__ import annotations

import importlib.util
import json
import re
import subprocess
import sys
from collections.abc import Iterable
from pathlib import Path

import pytest

SCRIPT_NAME = "check_pr_size.py"


def _find_script() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / "scripts" / SCRIPT_NAME
        if candidate.is_file():
            return candidate
    raise AssertionError(f"scripts/{SCRIPT_NAME} not found")


def _load():
    script = _find_script()
    spec = importlib.util.spec_from_file_location("check_pr_size", script)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_pr_size"] = module
    spec.loader.exec_module(module)
    return module


def _run(cwd: Path | None, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(_find_script()), *args],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
        cwd=str(cwd) if cwd is not None else None,
    )


def _payload(*, additions: int, deletions: int, labels: Iterable[str] = ()) -> dict:
    return {
        "pull_request": {
            "additions": additions,
            "deletions": deletions,
            "labels": [{"name": label} for label in labels],
        }
    }


@pytest.fixture
def tmp_git_repo(tmp_path: Path) -> Path:
    """A throwaway git repo on `main` with one empty initial commit."""
    subprocess.run(
        ["git", "init", "--initial-branch=main"],
        cwd=tmp_path,
        check=True,
        capture_output=True,
        encoding="utf-8",
    )
    subprocess.run(
        ["git", "config", "user.name", "Test"],
        cwd=tmp_path,
        check=True,
        capture_output=True,
        encoding="utf-8",
    )
    subprocess.run(
        ["git", "config", "user.email", "test@example.com"],
        cwd=tmp_path,
        check=True,
        capture_output=True,
        encoding="utf-8",
    )
    subprocess.run(
        ["git", "commit", "--allow-empty", "-m", "base"],
        cwd=tmp_path,
        check=True,
        capture_output=True,
        encoding="utf-8",
    )
    return tmp_path


def _commit(repo: Path, *, total: int) -> None:
    """Single commit on the branch adding `total` lines (insertions only)."""
    (repo / "new.txt").write_text("\n".join(["x"] * total) + "\n", encoding="utf-8")
    subprocess.run(
        ["git", "add", "new.txt"],
        cwd=repo,
        check=True,
        capture_output=True,
        encoding="utf-8",
    )
    subprocess.run(
        ["git", "commit", "-m", "feat"],
        cwd=repo,
        check=True,
        capture_output=True,
        encoding="utf-8",
    )


def _shortstat_total(repo: Path, base: str) -> int:
    """Read `git diff --shortstat` and return inserts + deletes."""
    raw = subprocess.run(
        ["git", "diff", "--shortstat", f"{base}...HEAD"],
        cwd=repo,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    ).stdout
    return _parse_shortstat_total(raw)


def _parse_shortstat_total(raw: str) -> int:
    """` 1 file changed, 250 insertions(+), 0 deletions(-)` -> 250."""
    inserts = 0
    deletes = 0
    for token in raw.split(","):
        token = token.strip()
        match = re.match(r"(\d+) insertion", token)
        if match:
            inserts = int(match.group(1))
        match = re.match(r"(\d+) deletion", token)
        if match:
            deletes = int(match.group(1))
    return inserts + deletes


# ---------------------------------------------------------------------------------------------
# Module-level invariants
# ---------------------------------------------------------------------------------------------


def test_ceiling_is_400() -> None:
    """The 400-line review budget is fixed."""
    module = _load()
    assert module.MAX_CHANGED_LINES == 400


# ---------------------------------------------------------------------------------------------
# Rule 1: ≤400 passes (exit 0).
# ---------------------------------------------------------------------------------------------


def test_local_under_budget_passes(tmp_git_repo: Path) -> None:
    _commit(tmp_git_repo, total=100)
    result = _run(tmp_git_repo, "local", "HEAD~1")
    assert result.returncode == 0, (
        f"under-budget local must pass; rc={result.returncode} stdout={result.stdout!r} "
        f"stderr={result.stderr!r}"
    )


def test_pr_event_under_budget_passes(tmp_path: Path) -> None:
    event = tmp_path / "event.json"
    event.write_text(json.dumps(_payload(additions=100, deletions=50)), encoding="utf-8")
    result = _run(None, "pr-event", str(event))
    assert result.returncode == 0, (
        f"under-budget pr-event must pass; rc={result.returncode} stdout={result.stdout!r} "
        f"stderr={result.stderr!r}"
    )


# ---------------------------------------------------------------------------------------------
# Rule 2: >400 without label fails (exit 1) with the total in the message.
# ---------------------------------------------------------------------------------------------


def test_pr_event_over_budget_without_label_fails(tmp_path: Path) -> None:
    event = tmp_path / "event.json"
    event.write_text(json.dumps(_payload(additions=350, deletions=200)), encoding="utf-8")
    result = _run(None, "pr-event", str(event))
    assert result.returncode == 1, (
        f"over-budget pr-event without label must fail; rc={result.returncode} "
        f"stdout={result.stdout!r} stderr={result.stderr!r}"
    )
    combined = result.stdout + result.stderr
    assert "550" in combined, f"total 550 must appear in output, got {combined!r}"


def test_local_over_budget_fails(tmp_git_repo: Path) -> None:
    _commit(tmp_git_repo, total=500)
    result = _run(tmp_git_repo, "local", "HEAD~1")
    assert result.returncode == 1, (
        f"over-budget local must fail (no override path); rc={result.returncode} "
        f"stdout={result.stdout!r} stderr={result.stderr!r}"
    )
    combined = result.stdout + result.stderr
    assert "500" in combined, f"total 500 must appear in output, got {combined!r}"


# ---------------------------------------------------------------------------------------------
# Rule 3: >400 with `size:exception` label passes.
# ---------------------------------------------------------------------------------------------


def test_pr_event_over_budget_with_size_exception_label_passes(tmp_path: Path) -> None:
    event = tmp_path / "event.json"
    event.write_text(
        json.dumps(_payload(additions=350, deletions=200, labels=["size:exception"])),
        encoding="utf-8",
    )
    result = _run(None, "pr-event", str(event))
    assert result.returncode == 0, (
        f"`size:exception` label must override over-budget; rc={result.returncode} "
        f"stdout={result.stdout!r} stderr={result.stderr!r}"
    )


# ---------------------------------------------------------------------------------------------
# Rule 4: `local` gives the same total as the API for the same range.
# ---------------------------------------------------------------------------------------------


def test_local_total_matches_shortstat(tmp_git_repo: Path) -> None:
    """`local` parses `git diff --shortstat` so it matches the API total."""
    _commit(tmp_git_repo, total=250)
    expected = _shortstat_total(tmp_git_repo, "HEAD~1")
    assert expected == 250

    module = _load()
    total = module.compute_local_total(cwd=tmp_git_repo, merge_base="HEAD~1")
    assert total == 250


# ---------------------------------------------------------------------------------------------
# Rule 5: binary output, no warning.
# ---------------------------------------------------------------------------------------------


def test_no_warning_on_pr_event_pass(tmp_path: Path) -> None:
    event = tmp_path / "event.json"
    event.write_text(json.dumps(_payload(additions=10, deletions=5)), encoding="utf-8")
    result = _run(None, "pr-event", str(event))
    assert result.stderr == "", f"pr-event pass must be silent on stderr, got {result.stderr!r}"


def test_no_warning_on_local_pass(tmp_git_repo: Path) -> None:
    _commit(tmp_git_repo, total=10)
    result = _run(tmp_git_repo, "local", "HEAD~1")
    assert result.stderr == "", f"local pass must be silent on stderr, got {result.stderr!r}"


def test_no_warning_on_pr_event_fail(tmp_path: Path) -> None:
    """A fail prints the verdict on stderr (single concise message), nothing on stdout."""
    event = tmp_path / "event.json"
    event.write_text(json.dumps(_payload(additions=500, deletions=10)), encoding="utf-8")
    result = _run(None, "pr-event", str(event))
    assert result.returncode == 1
    assert "510" in (result.stdout + result.stderr)
    # Python deprecation/runtime warnings would land on stderr — refuse them.
    assert "DeprecationWarning" not in result.stderr
    assert "UserWarning" not in result.stderr
