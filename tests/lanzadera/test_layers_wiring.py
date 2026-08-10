# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp DA-1 — test_layers_wiring.py
"""Wiring pin for `scripts/check_layers.py` (QC-2 / QC-9, DA-1).

The hexagonal layer gate enforces dependency direction, vertical slicing, and
purity inside `platform.src.modules`. Phase 0 asserts the gate is wired,
configures the correct `ROOT_PACKAGE`, and returns a valid JSON envelope on an
empty package.
"""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from pathlib import Path

import pytest  # noqa: F401 — module imported for fixture decorators


def _find_script() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / "scripts" / "check_layers.py"
        if candidate.is_file():
            return candidate
    raise AssertionError("scripts/check_layers.py not found")


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


def test_script_is_importable(script: Path) -> None:
    """Phase 0 RED-then-GREEN: the gate module must be loadable as Python."""
    spec = importlib.util.spec_from_file_location("check_layers", script)
    assert spec and spec.loader, f"could not load {script}"
    module = importlib.util.module_from_spec(spec)
    # Register in sys.modules BEFORE exec so dataclass decorators can resolve
    # `cls.__module__` against the actual module dict. Without this, Python 3.13
    # raises `AttributeError: 'NoneType' object has no attribute '__dict__'` when
    # the dataclass tries to resolve a `from __future__ import annotations`
    # string.
    sys.modules["check_layers"] = module
    spec.loader.exec_module(module)
    assert module.ROOT_PACKAGE == "app.src.modules"


def test_layers_gate_runs_clean_on_empty_package(root: Path, script: Path) -> None:
    """Phase 0 wiring: an empty `app.src.modules` must report `OK layer gate clean`."""
    result = subprocess.run(
        [sys.executable, str(script), "--root", str(root)],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 0, (
        f"check_layers.py exited {result.returncode}\nSTDOUT: {result.stdout}\n"
        f"STDERR: {result.stderr}"
    )
    assert "OK    layer gate clean" in result.stdout


def test_layers_gate_emits_valid_envelope(root: Path, script: Path) -> None:
    """The JSON envelope schema is the contract — indicators, ceilings, findings."""
    result = subprocess.run(
        [sys.executable, str(script), "--root", str(root), "--json"],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 0, (
        f"check_layers.py exited {result.returncode}\nSTDERR: {result.stderr}"
    )
    envelope = json.loads(result.stdout)
    assert envelope["gate"] == "layers"
    assert envelope["status"] == "pass"
    assert "indicators" in envelope
    assert "files_checked" in envelope["indicators"]
    assert "ceilings" in envelope
    assert "findings" in envelope
    assert envelope["indicators"]["files_checked"] == 0
    assert envelope["indicators"]["violations"] == 0


def test_layers_gate_fails_closed_when_root_package_missing(
    tmp_path: Path, script: Path
) -> None:
    """Hard Rule 18: missing root package is an error, not a silent pass."""
    result = subprocess.run(
        [sys.executable, str(script), "--root", str(tmp_path), "--json"],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    assert result.returncode == 1, "missing root package must exit non-zero"
    envelope = json.loads(result.stdout)
    assert envelope["gate"] == "layers"
    assert envelope["status"] == "error"