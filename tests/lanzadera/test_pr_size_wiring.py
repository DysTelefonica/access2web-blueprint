# HARNESS-PROVENANCE: access2web-blueprint issue #854 — test_pr_size_wiring.py
"""Wiring pin for `scripts/check_pr_size.py` after the #854 refactor.

Counts additions + deletions and fails above 400 (gentle-ai reference,
`.github/scripts/check-pr-size.cjs:106-110`). Override is the
`size:exception` label only — no body reason line. One parametrized test
per acceptance rule.
"""

from __future__ import annotations

import importlib.util
import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[2] / "scripts" / "check_pr_size.py"


def _load():
    spec = importlib.util.spec_from_file_location("check_pr_size", SCRIPT)
    m = importlib.util.module_from_spec(spec)
    sys.modules["check_pr_size"] = m
    spec.loader.exec_module(m)
    return m


def _run(working_dir, *args):
    return subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        cwd=str(working_dir) if working_dir else None,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )


def _payload(additions, deletions, *labels):
    return {
        "pull_request": {
            "additions": additions,
            "deletions": deletions,
            "labels": [{"name": label} for label in labels],
        }
    }


def _event_with(tmp_path, additions, deletions, *labels):
    path = tmp_path / "event.json"
    path.write_text(json.dumps(_payload(additions, deletions, *labels)), encoding="utf-8")
    return str(path)


def _local_repo_with_diff(tmp_path, total):
    """Init a `main`-first repo, add one commit of `total` insertions; return its path."""
    repo = tmp_path / "local"
    repo.mkdir()
    subprocess.run(
        ["git", "init", "--initial-branch=main"],
        cwd=repo,
        check=True,
        capture_output=True,
    )
    for k, v in [("user.email", "t@e"), ("user.name", "t")]:
        subprocess.run(["git", "config", k, v], cwd=repo, check=True, capture_output=True)
    subprocess.run(
        ["git", "commit", "--allow-empty", "-m", "b"],
        cwd=repo,
        check=True,
        capture_output=True,
    )
    (repo / "f.txt").write_text("\n".join(["x"] * total) + "\n", encoding="utf-8")
    subprocess.run(["git", "add", "f.txt"], cwd=repo, check=True, capture_output=True)
    subprocess.run(["git", "commit", "-m", "f"], cwd=repo, check=True, capture_output=True)
    return repo


def test_ceiling_is_400():
    """The 400-line review budget is fixed."""
    assert _load().MAX_CHANGED_LINES == 400


# Rules 1-3: the verdict matrix — (mode, adds, dels, labels) → exit code (& total in stderr
# when failing). Local mode has no label path; Rule 3 only applies to pr-event.
@pytest.mark.parametrize(
    ("mode", "adds", "dels", "labels", "expected_rc", "expected_in_output"),
    [
        # Rule 1: ≤400 → exit 0 in both modes.
        ("pr-event", 100, 50, (), 0, None),
        ("local", 100, 0, (), 0, None),
        # Rule 2: >400 without label → exit 1, total echoed.
        ("pr-event", 350, 200, (), 1, "550"),
        ("local", 500, 0, (), 1, "500"),
        # Rule 3: >400 with `size:exception` → exit 0 (pr-event only).
        ("pr-event", 350, 200, ("size:exception",), 0, None),
    ],
    ids=["r1-pr-event", "r1-local", "r2-pr-event", "r2-local", "r3-label"],
)
def test_rules_1_2_3_verdict(mode, adds, dels, labels, expected_rc, expected_in_output, tmp_path):
    if mode == "pr-event":
        event = _event_with(tmp_path, adds, dels, *labels)
        result = _run(None, mode, event)
    else:
        repo = _local_repo_with_diff(tmp_path, adds + dels)
        result = _run(repo, mode, "HEAD~1")
    assert result.returncode == expected_rc, (
        f"unexpected exit: scenario={mode}+{adds}+{dels}+{labels} "
        f"got {result.returncode}; stderr={result.stderr!r}"
    )
    if expected_in_output:
        assert expected_in_output in (result.stdout + result.stderr), (
            f"{expected_in_output!r} must appear in output: "
            f"stdout={result.stdout!r}, stderr={result.stderr!r}"
        )


# Rule 4: `local` parses `git diff --shortstat`, equal to the API total for the same range.
def test_rule_4_local_matches_shortstat(tmp_path):
    total = 250
    repo = _local_repo_with_diff(tmp_path, total)
    shortstat = subprocess.run(
        ["git", "diff", "--shortstat", "HEAD~1...HEAD"],
        cwd=repo,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    ).stdout
    api_total = sum(
        int(m.group(1)) for m in re.finditer(r"(\d+) (?:insertion|deletion)", shortstat)
    )
    assert api_total == total  # sanity: shortstat equals what we'd feed to the API.
    assert _load().compute_local_total(cwd=repo, merge_base="HEAD~1") == api_total


# Rule 5: binary output, no Python warnings on stderr.
@pytest.mark.parametrize("verdict_path", ["pass", "fail"], ids=["pass", "fail"])
def test_rule_5_binary_output_no_warning(verdict_path, tmp_path):
    if verdict_path == "pass":
        event = _event_with(tmp_path, 10, 5)
        result = _run(None, "pr-event", event)
        assert result.returncode == 0
    else:
        event = _event_with(tmp_path, 500, 10)
        result = _run(None, "pr-event", event)
        assert result.returncode == 1
        assert "510" in (result.stdout + result.stderr)
    for warning in ("DeprecationWarning", "UserWarning"):
        assert warning not in result.stderr, f"unexpected {warning} on stderr: {result.stderr!r}"
