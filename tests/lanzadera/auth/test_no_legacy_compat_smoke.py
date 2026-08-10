# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp DA-13 — test_no_legacy_compat_smoke.py
"""Smoke import for the DA-13 pin-test harness (lanzadera-mvp Phase 0).

Phase 0 wires the AST walker (`scripts/check_legacy_hashes.py`) but does NOT
ship the full pytest integration — the AST walker that fails the suite on
`legacy_hash`, `verify_legacy`, `sha256`, `old_password`, `migrate_password`
lives at `tests/lanzadera/auth/test_no_legacy_compat.py` in Phase 6.

This smoke test asserts the walker is importable, exposes the forbidden
symbol set, and passes on the current source tree (no legacy symbols yet).
If a future PR introduces one of those names anywhere under `platform/src/`,
the walker would surface it; this test makes the absence a single line of
code.
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
        candidate = parent / "scripts" / "check_legacy_hashes.py"
        if candidate.is_file():
            return candidate
    raise AssertionError("scripts/check_legacy_hashes.py not found")


def _find_root() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / "app" / "src"
        if candidate.is_dir():
            return parent
    raise AssertionError("worktree root not found (no app/src/)")


@pytest.fixture(scope="module")
def script() -> Path:
    return _find_script()


@pytest.fixture(scope="module")
def root() -> Path:
    return _find_root()


def test_da13_walker_is_importable(script: Path) -> None:
    """RED-then-GREEN: the script MUST load as a Python module.

    A broken import (e.g. syntax error in the walker) is a hard fail. The
    walker cannot be enforced from pytest if it cannot be imported at all.
    """
    spec = importlib.util.spec_from_file_location("check_legacy_hashes", script)
    assert spec and spec.loader, f"could not load {script}"
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_legacy_hashes"] = module
    spec.loader.exec_module(module)


def test_da13_forbidden_symbols_declared(script: Path) -> None:
    """The five forbidden symbols are the contract."""
    spec = importlib.util.spec_from_file_location("check_legacy_hashes", script)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_legacy_hashes"] = module
    spec.loader.exec_module(module)
    assert "legacy_hash" in module.FORBIDDEN_SYMBOLS
    assert "verify_legacy" in module.FORBIDDEN_SYMBOLS
    assert "sha256" in module.FORBIDDEN_SYMBOLS
    assert "old_password" in module.FORBIDDEN_SYMBOLS
    assert "migrate_password" in module.FORBIDDEN_SYMBOLS


def test_da13_walker_passes_on_empty_app_src(root: Path, script: Path) -> None:
    """Phase 0 wiring: no legacy symbols yet, walker must exit 0."""
    result = subprocess.run(
        [sys.executable, str(script), "--root", str(root)],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 0, (
        f"check_legacy_hashes.py must pass on the current source tree\n"
        f"STDOUT: {result.stdout}\nSTDERR: {result.stderr}"
    )
    assert "OK    no legacy crypto symbols in app/src/" in result.stdout