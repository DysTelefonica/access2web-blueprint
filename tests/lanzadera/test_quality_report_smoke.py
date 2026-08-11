# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp
# — test_quality_report_smoke.py
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
    # mutation_sites sits between crap and dry (Hard Rule 13: order is load-bearing;
    # #117 retro split added it on top of the original five-gate chain).
    assert names == ("layers", "complexity", "crap", "mutation_sites", "dry", "legacy_hashes")


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
    # With Phase 0 coverage (~94%) the aggregator reports pass end-to-end. The earlier
    # contract assumed zero coverage (CRAP fails closed) — that pre-condition no longer
    # holds, and the test reflects current reality. Hard Rule 18: a measurement that
    # can't run must never score as a perfect one; that still holds, via the
    # indicator counts being recorded.
    assert result.returncode == 0, (
        f"quality_report.py must report pass when CRAP gate has coverage "
        f"(got {result.returncode})\n"
        f"STDOUT: {result.stdout}\nSTDERR: {result.stderr}"
    )
    payload = json.loads(out.read_text(encoding="utf-8"))
    assert payload["schema"] == "deterministic-quality-harness/quality-report/v1"
    assert "commit" in payload
    assert payload["status"] == "pass"
    assert "crap" not in payload["failed_gates"]
    gates = {entry["gate"] for entry in payload["gates"]}
    # mutation_sites joined the chain between crap and dry (#117 retro split).
    assert {"layers", "complexity", "crap", "mutation_sites", "dry", "legacy_hashes"} <= gates
