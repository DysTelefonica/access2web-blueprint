# HARNESS-PROVENANCE: access2web-blueprint issue #854 — test_pr_size_wiring.py
"""Wiring pin for `scripts/check_pr_size.py` (issue #854): one test per acceptance rule.

The budget is additions + deletions, at most 400. The only override is the
`size:exception` label. Rules 1-3 share one parametrized verdict matrix.
"""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[2] / "scripts" / "check_pr_size.py"


def _load():
    spec = importlib.util.spec_from_file_location("check_pr_size", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_pr_size"] = module
    spec.loader.exec_module(module)
    return module


def _run(*args, cwd=None):
    return subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        cwd=cwd,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )


def _git(repo, *args):
    return subprocess.run(
        ["git", *args], cwd=repo, check=True, capture_output=True, text=True
    ).stdout


def _event(tmp_path, additions, deletions, labels=()):
    payload = {
        "pull_request": {
            "additions": additions,
            "deletions": deletions,
            "labels": [{"name": name} for name in labels],
        }
    }
    path = tmp_path / "event.json"
    path.write_text(json.dumps(payload), encoding="utf-8")
    return str(path)


def _local_repo(tmp_path, insertions, deletions):
    """Repo whose HEAD replaces the `deletions` lines of HEAD~1 with `insertions` new lines."""
    repo = tmp_path / "local"
    repo.mkdir()
    _git(repo, "init", "--initial-branch=main")
    _git(repo, "config", "user.email", "t@e")
    _git(repo, "config", "user.name", "t")
    (repo / "f.txt").write_text("old\n" * deletions, encoding="utf-8")
    _git(repo, "add", "f.txt")
    _git(repo, "commit", "-m", "base")
    (repo / "f.txt").write_text("new\n" * insertions, encoding="utf-8")
    _git(repo, "commit", "-am", "change")
    return repo


def test_ceiling_is_400():
    """The 400-line review budget is fixed."""
    assert _load().MAX_CHANGED_LINES == 400


# Rules 1-3: (mode, additions or insertions, deletions, labels) -> exit code, and the total
# in the output when over budget.
@pytest.mark.parametrize(
    ("mode", "adds", "dels", "labels", "expected_rc", "expected_total"),
    [
        ("pr-event", 100, 50, (), 0, None),  # rule 1: within budget
        ("local", 100, 0, (), 0, None),
        ("pr-event", 350, 200, (), 1, "550"),  # rule 2: over budget, no label
        ("local", 500, 0, (), 1, "500"),
        ("pr-event", 350, 200, ("size:exception",), 0, None),  # rule 3: the label overrides
    ],
    ids=["r1-pr-event", "r1-local", "r2-pr-event", "r2-local", "r3-label"],
)
def test_rules_1_2_3_verdict(mode, adds, dels, labels, expected_rc, expected_total, tmp_path):
    if mode == "pr-event":
        result = _run(mode, _event(tmp_path, adds, dels, labels))
    else:
        result = _run(mode, "HEAD~1", cwd=_local_repo(tmp_path, adds, dels))
    assert result.returncode == expected_rc, result.stderr
    if expected_total:
        assert expected_total in result.stderr


def test_rule_4_local_total_equals_the_api_total(tmp_path):
    """`local` gives the total GitHub reports for the same range: 200 added + 50 removed."""
    repo = _local_repo(tmp_path, insertions=200, deletions=50)
    assert _load().compute_local_total(repo, "HEAD~1") == 250


@pytest.mark.parametrize(
    ("adds", "dels", "expected_rc"),
    [(10, 5, 0), (500, 10, 1)],
    ids=["pass", "fail"],
)
def test_rule_5_binary_output_without_warnings(adds, dels, expected_rc, tmp_path):
    """The verdict is the exit code; stdout stays empty and stderr carries no Python warning."""
    result = _run("pr-event", _event(tmp_path, adds, dels))
    assert result.returncode == expected_rc
    assert result.stdout == ""
    assert "Warning" not in result.stderr
