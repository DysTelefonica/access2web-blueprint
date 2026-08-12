# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 — assets/tests/test_gate_liveness.py
"""Hard Rule 18 for the tree-walking gates: an empty subject set is not a pass.

Every gate here already refused to run when its root package was *missing*. None
refused when the package existed and held nothing to inspect — and that is the
shape the rule warns about, because it looks like success rather than an error.
A complexity gate over zero functions reports "every function at or below 15"; a
DRY gate over zero files reports "no duplicated block". Both are the healthiest
possible number for the least healthy possible state (issue #143).

The clean/violating cases are the regression half: the guard must not soften a
real verdict, or it would be "fixed" later by making the gate quieter.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = REPO_ROOT / "scripts"
FIXTURES = Path(__file__).resolve().parent / "fixtures"

#: The gates that walk a source tree. `check_crap` is deliberately absent: it
#: consumes coverage.json and already fails closed when that file is missing.
TREE_GATES = ("check_complexity", "check_dry", "check_mutation_sites", "check_layers")


def _run(gate: str, root: Path) -> int:
    return subprocess.run(
        [sys.executable, str(SCRIPTS / f"{gate}.py"), "--root", str(root)],
        capture_output=True,
        text=True,
    ).returncode


@pytest.mark.parametrize("gate", TREE_GATES)
def test_empty_subject_set_fails_closed(gate: str, tmp_path: Path) -> None:
    """The package exists and is empty — the silent case, not the missing-root one."""
    (tmp_path / "app").mkdir()

    assert _run(gate, tmp_path) == 1, (
        f"{gate} scored an empty tree as clean. A measurement that could not run "
        f"must never report the healthiest possible number (Hard Rule 18)."
    )


@pytest.mark.parametrize("gate", TREE_GATES)
def test_missing_root_package_still_fails(gate: str, tmp_path: Path) -> None:
    """The guard that already existed must survive the new one."""
    assert _run(gate, tmp_path) == 1


@pytest.mark.parametrize("gate", TREE_GATES)
def test_clean_fixture_still_passes(gate: str) -> None:
    """Regression: the liveness guard must not turn a real pass into a failure."""
    assert _run(gate, FIXTURES / "complexity_clean") == 0


def test_violating_fixture_still_fails() -> None:
    """Regression the other way: the gate is not softened by the guard."""
    assert _run("check_complexity", FIXTURES / "complexity_violation") == 1
    assert _run("check_dry", FIXTURES / "dry_violation") == 1
