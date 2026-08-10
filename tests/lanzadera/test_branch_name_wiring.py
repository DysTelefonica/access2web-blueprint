# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp QC-6 — test_branch_name_wiring.py
"""Wiring pin for `scripts/check_branch_name.py` (QC-6).

The branch-name gate accepts branches matching the convention
`^(feat|fix|refactor|docs|ci|test|chore)/<n>-<slug>$`, plus an allowlist for
long-lived branches. Phase 0 asserts the gate is wired, that the regex is
intact, and that the tracker branch `feat/lanzadera-mvp-tracker` is allowed
because it aggregates the feature-branch-chain (decided 2026-08-09).
"""

from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path

import pytest


def _find_script() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / "scripts" / "check_branch_name.py"
        if candidate.is_file():
            return candidate
    raise AssertionError("scripts/check_branch_name.py not found")


@pytest.fixture(scope="module")
def script() -> Path:
    return _find_script()


def test_branch_name_pattern_is_intact(script: Path) -> None:
    spec = importlib.util.spec_from_file_location("check_branch_name", script)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_branch_name"] = module
    spec.loader.exec_module(module)
    assert module.PATTERN.pattern == (
        r"^(feat|fix|refactor|docs|ci|test|chore)/\d+-[a-z0-9]+(-[a-z0-9]+)*$"
    )


def test_tracker_branch_is_allowlisted(script: Path) -> None:
    spec = importlib.util.spec_from_file_location("check_branch_name", script)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_branch_name"] = module
    spec.loader.exec_module(module)
    # The chain strategy documents that `feat/lanzadera-mvp-tracker` aggregates
    # PR 1..PR 8; the allowlist MUST include it.
    assert "feat/lanzadera-mvp-tracker" in module.ALLOWLIST


def test_branch_name_gate_accepts_conventional_branches(script: Path) -> None:
    result = subprocess.run(
        [sys.executable, str(script), "--branch", "feat/1234-lanzadera-auth"],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 0, (
        f"conventional branch must pass (got {result.returncode})\n"
        f"STDOUT: {result.stdout}\nSTDERR: {result.stderr}"
    )
    assert "matches the convention" in result.stdout


def test_branch_name_gate_rejects_non_conventional_branches(script: Path) -> None:
    result = subprocess.run(
        [sys.executable, str(script), "--branch", "no-type-prefix"],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 1, "non-conventional branch must fail the gate"


def test_branch_name_gate_accepts_allowlisted_branches(script: Path) -> None:
    result = subprocess.run(
        [sys.executable, str(script), "--branch", "feat/lanzadera-mvp-tracker"],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 0, (
        f"tracker branch must be allowlisted (got {result.returncode})\n{result.stdout}"
    )
    assert "allowlisted" in result.stdout