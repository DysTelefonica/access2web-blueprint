# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp — test_quality_report_smoke.py
"""Wiring pin for `scripts/quality_report.py` (QC-11).

The aggregator runs every gate in a fixed order and merges the envelopes.
Phase 0 asserts the aggregator runs, produces `quality-report.json`, and that
the gate order is the contract (`layers -> complexity -> crap -> dry ->
legacy_hashes`).
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
        candidate = parent / "scripts" / "quality_report.py"
        if candidate.is_file():
            return candidate
    raise AssertionError("scripts/quality_report.py not found")


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


def test_gate_order_includes_legacy_hashes(script: Path) -> None:
    """Hard Rule 13 + DA-13: the gate order in code, not in YAML."""
    spec = importlib.util.spec_from_file_location("quality_report", script)
    module = importlib.util.module_from_spec(spec)
    sys.modules["quality_report"] = module
    spec.loader.exec_module(module)
    names = tuple(name for name, _, _ in module.GATES)
    assert names == ("layers", "complexity", "crap", "dry", "legacy_hashes")


def test_quality_report_produces_json_envelope(tmp_path: Path, root: Path, script: Path) -> None:
    """The aggregator runs every gate and writes `quality-report.json`."""
    out = tmp_path / "quality-report.json"
    result = subprocess.run(
        [
            sys.executable,
            str(script),
            "--root",
            str(root),
            "--out",
            str(out),
        ],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    # With Phase 0 source, layers and complexity pass; CRAP fails closed (no coverage).
    # The aggregator reports `fail` (because CRAP is in the chain) — that's the contract.
    assert result.returncode == 1, (
        f"quality_report.py must report fail when CRAP gate has no coverage (got {result.returncode})\n"
        f"STDOUT: {result.stdout}\nSTDERR: {result.stderr}"
    )
    payload = json.loads(out.read_text(encoding="utf-8"))
    assert payload["schema"] == "deterministic-quality-harness/quality-report/v1"
    assert "commit" in payload
    assert payload["status"] == "fail"
    assert "crap" in payload["failed_gates"]
    gates = {entry["gate"] for entry in payload["gates"]}
    assert {"layers", "complexity", "crap", "dry", "legacy_hashes"} <= gates