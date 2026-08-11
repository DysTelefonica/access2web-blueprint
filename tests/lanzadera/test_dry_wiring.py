# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp QC-11 — test_dry_wiring.py
"""Wiring pin for `scripts/check_dry.py` (QC-11).

The DRY gate detects type-1 and type-2 duplicate blocks of `MIN_STATEMENTS = 5`
consecutive statements on a normalised AST. Phase 0 asserts the gate is wired
and reports zero duplicates on an empty package.
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
        candidate = parent / "scripts" / "check_dry.py"
        if candidate.is_file():
            return candidate
    raise AssertionError("scripts/check_dry.py not found")


def _find_root() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / "app" / "src" / "modules"
        if candidate.is_dir():
            return parent
    raise AssertionError("worktree root not found (no app/src/modules/)")


@pytest.fixture(scope="module")
def root() -> Path:
    return _find_root()


@pytest.fixture(scope="module")
def script() -> Path:
    return _find_script()


def test_dry_min_statements_is_five(script: Path) -> None:
    spec = importlib.util.spec_from_file_location("check_dry", script)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_dry"] = module
    spec.loader.exec_module(module)
    assert module.MIN_STATEMENTS == 5
    assert module.MIN_OCCURRENCES == 2


def test_dry_gate_runs_clean_on_empty_package(root: Path, script: Path) -> None:
    result = subprocess.run(
        [sys.executable, str(script), "--root", str(root)],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 0, (
        f"check_dry.py exited {result.returncode}\nSTDOUT: {result.stdout}\nSTDERR: {result.stderr}"
    )
    assert "OK    no duplicated block of 5+ statements" in result.stdout


def test_dry_gate_emits_valid_envelope(root: Path, script: Path) -> None:
    result = subprocess.run(
        [sys.executable, str(script), "--root", str(root), "--json"],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    envelope = json.loads(result.stdout)
    assert envelope["gate"] == "dry"
    assert envelope["status"] == "pass"
    assert envelope["indicators"]["duplicate_groups"] == 0
