# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp QC-11 — test_crap_wiring.py
"""Wiring pin for `scripts/check_crap.py` (QC-11).

The CRAP gate (`CRAP <= 6`) prices complexity by how well it is tested. It
fails closed: no coverage data means no verdict (Hard Rule 18). Phase 0
asserts the gate is wired and emits an `error` envelope when coverage is
absent — it must never silently report success.
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
        candidate = parent / "scripts" / "check_crap.py"
        if candidate.is_file():
            return candidate
    raise AssertionError("scripts/check_crap.py not found")


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


def test_crap_ceiling_is_six(script: Path) -> None:
    spec = importlib.util.spec_from_file_location("check_crap", script)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_crap"] = module
    spec.loader.exec_module(module)
    assert module.MAX_CRAP == 6.0


def test_crap_gate_fails_closed_without_coverage(root: Path, script: Path) -> None:
    """Hard Rule 18: a missing coverage file is an error, not a silent pass."""
    fake_coverage = root / "coverage.missing.json"
    if fake_coverage.exists():
        fake_coverage.unlink()
    result = subprocess.run(
        [
            sys.executable,
            str(script),
            "--root",
            str(root),
            "--coverage-json",
            str(fake_coverage),
            "--json",
        ],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 1, (
        f"check_crap.py must exit 1 when coverage is missing (got {result.returncode})\n"
        f"STDOUT: {result.stdout}\nSTDERR: {result.stderr}"
    )
    envelope = json.loads(result.stdout)
    assert envelope["gate"] == "crap"
    assert envelope["status"] == "error"
    assert "not found" in envelope["detail"]
