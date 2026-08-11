# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp QC-6
# — test_pr_size_wiring.py
"""Wiring pin for `scripts/check_pr_size.py` (QC-6).

The PR size gate counts changed lines against a base ref and fails above 400.
Phase 0 asserts the gate parses its arguments, evaluates an override when one
is supplied with a reason, and emits a JSON envelope with the right schema.
"""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from pathlib import Path

import pytest


def _find_script() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / "scripts" / "check_pr_size.py"
        if candidate.is_file():
            return candidate
    raise AssertionError("scripts/check_pr_size.py not found")


def _find_root() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        # In a worktree, `.git` is a FILE pointing to the real gitdir; in a
        # full clone, it is a DIRECTORY. Accept both.
        candidate = parent / ".git"
        if candidate.is_dir() or candidate.is_file():
            return parent
    raise AssertionError("worktree root not found (no .git)")


@pytest.fixture(scope="module")
def root() -> Path:
    return _find_root()


@pytest.fixture(scope="module")
def script() -> Path:
    return _find_script()


def test_pr_size_ceiling_is_400(script: Path) -> None:
    spec = importlib.util.spec_from_file_location("check_pr_size", script)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_pr_size"] = module
    spec.loader.exec_module(module)
    assert module.MAX_CHANGED_LINES == 400


def test_pr_size_override_needs_reason() -> None:
    """Hard Rule: the override is awkward on purpose — `size:exception` alone does nothing."""
    spec = importlib.util.spec_from_file_location("check_pr_size", _find_script())
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_pr_size"] = module
    spec.loader.exec_module(module)
    # No marker, no reason — both should be None.
    assert module.override_reason("plain PR body without marker") is None
    # Marker but no reason — None.
    assert module.override_reason("size:exception\nno reason line here\n") is None
    # Marker AND reason — the reason line wins.
    body = "size:exception\nsize-exception-reason: PR 1 is the indivisible foundation scaffold.\n"
    assert module.override_reason(body) == "PR 1 is the indivisible foundation scaffold."


def test_pr_size_gate_runs_against_head(root: Path, script: Path) -> None:
    """Phase 0 wiring: the gate runs end-to-end against the local commit chain."""
    result = subprocess.run(
        [
            sys.executable,
            str(script),
            "--base-ref",
            "HEAD~0",
            "--json",
        ],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
        cwd=str(root),
    )
    # The gate may report `error` when the diff is empty (HEAD~0 is HEAD, no diff),
    # but the JSON envelope must always parse.
    assert result.returncode in (0, 1), f"unexpected exit {result.returncode}"
    envelope = json.loads(result.stdout)
    assert envelope["gate"] == "pr_size"
    assert "indicators" in envelope
    assert "changed_lines" in envelope["indicators"]
    assert envelope["ceilings"]["changed_lines"] == 400
