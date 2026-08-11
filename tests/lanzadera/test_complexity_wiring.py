# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp QC-10 — test_complexity_wiring.py
"""Wiring pin for `scripts/check_complexity.py` (QC-10).

The complexity gate enforces an absolute, global ceiling of `CC <= 15` on every
function. Phase 0 asserts the gate is wired and returns a valid JSON envelope
on an empty package.
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
        candidate = parent / "scripts" / "check_complexity.py"
        if candidate.is_file():
            return candidate
    raise AssertionError("scripts/check_complexity.py not found")


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


def test_max_complexity_ceiling_is_absolute(script: Path) -> None:
    """Hard Rule 12: a global ceiling, never `top-N`."""
    spec = importlib.util.spec_from_file_location("check_complexity", script)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_complexity"] = module
    spec.loader.exec_module(module)
    assert module.MAX_COMPLEXITY == 15
    assert module.BASELINE == {}, "Phase 0 ships with an empty BASELINE — greenfield"


def test_complexity_gate_runs_clean_on_empty_package(root: Path, script: Path) -> None:
    """Phase 0 wiring: no source files yet means no offenders and no functions measured."""
    result = subprocess.run(
        [sys.executable, str(script), "--root", str(root)],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 0, (
        f"check_complexity.py exited {result.returncode}\nSTDOUT: {result.stdout}\n"
        f"STDERR: {result.stderr}"
    )
    assert "OK    every function at or below complexity 15" in result.stdout


def test_complexity_gate_emits_valid_envelope(root: Path, script: Path) -> None:
    result = subprocess.run(
        [sys.executable, str(script), "--root", str(root), "--json"],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 0, f"check_complexity.py exited {result.returncode}"
    envelope = json.loads(result.stdout)
    assert envelope["gate"] == "complexity"
    assert envelope["status"] == "pass"
    assert envelope["indicators"]["functions_over_ceiling"] == 0
    # The Phase 0 composition root (`app/src/main.py`) contributes a couple
    # of trivial functions; `max_complexity` is the high-water mark and MUST
    # stay at or below 15. We assert the ceiling rather than the exact value
    # so the test is robust against new trivial helpers in Phase 1+.
    assert envelope["indicators"]["max_complexity"] <= 15
    assert envelope["ceilings"]["max_complexity"] == 15
