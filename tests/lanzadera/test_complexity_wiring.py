# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp QC-10
# + architectural-guards-over-metrics DG-11 — test_complexity_wiring.py
"""Wiring pin for `scripts/check_complexity.py` (QC-10, DG-11).

The complexity gate enforces an absolute, global ceiling of `CC <= 10` on every
function. The ceiling is reviewed on `2027-02-13`: if the code stays at CC 6,
the ceiling moves toward 6. Phase 0 asserts the gate is wired and returns a
valid JSON envelope; Phase 1 (issue #266) lowers the ceiling from 15 to 10 and
pins the review date next to the constant.
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
    """Hard Rule 12: a global ceiling, never `top-N`.

    DG-11 (issue #266) lowers the ceiling from 15 to 10 with a review date of
    2027-02-13. The module source MUST declare both the ceiling and the date.
    """
    spec = importlib.util.spec_from_file_location("check_complexity", script)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_complexity"] = module
    spec.loader.exec_module(module)
    assert module.MAX_COMPLEXITY == 10
    assert module.BASELINE == {}, "Phase 0 ships with an empty BASELINE — greenfield"


def test_max_complexity_source_declares_review_date(script: Path) -> None:
    """DG-11: the review date is written next to the ceiling constant.

    The date is grepped from the source so that a future edit cannot lower the
    constant without keeping the date — and a future edit cannot erase the date
    without lowering the constant back.
    """
    source = script.read_text(encoding="utf-8")
    assert "2027-02-13" in source, (
        "DG-11 review date missing from scripts/check_complexity.py source"
    )
    # The date MUST be in the contiguous comment block immediately preceding
    # the MAX_COMPLEXITY assignment. "Adjacent" here means "in the same docstring
    # the constant belongs to" — a casual reader must see them together.
    lines = source.splitlines()
    max_complexity_lines = [
        index for index, line in enumerate(lines) if "MAX_COMPLEXITY" in line and "=" in line
    ]
    assert max_complexity_lines, "MAX_COMPLEXITY assignment not found in source"
    target_index = max_complexity_lines[0]
    # Walk backwards from the constant through the contiguous `#:` comment
    # block (Python docstring convention for module-level variables). Stop at
    # the first non-comment line.
    preceding_comment_block: list[str] = []
    for index in range(target_index - 1, -1, -1):
        line = lines[index]
        if line.lstrip().startswith("#"):
            preceding_comment_block.append(line)
        else:
            break
    assert any("2027-02-13" in line for line in preceding_comment_block), (
        "DG-11: 2027-02-13 review date must appear in the contiguous comment "
        "block immediately preceding MAX_COMPLEXITY; got: "
        f"{preceding_comment_block!r}"
    )


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
    assert "OK    every function at or below complexity 10" in result.stdout


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
    # stay at or below the current ceiling. We assert the ceiling rather than
    # the exact value so the test is robust against new trivial helpers.
    assert envelope["indicators"]["max_complexity"] <= 10
    assert envelope["ceilings"]["max_complexity"] == 10
