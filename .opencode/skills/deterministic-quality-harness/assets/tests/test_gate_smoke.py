# HARNESS-PROVENANCE: deterministic-quality-harness v1.2 — assets/tests/test_gate_smoke.py
"""Smoke tests: every gate is observed exiting 1 on a real violation (Execution Step 5).

A gate that has never been seen failing is a false guarantee. Each test here feeds a known
violation to a gate and asserts the failure, then feeds clean input and asserts the pass. When a
refactor silently breaks a detector, these go red before the detector's absence does damage.
"""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
FIXTURES = Path(__file__).resolve().parent / "fixtures"


def _run(script: str, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPTS / script), *args],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )


def _load(script: str):
    spec = importlib.util.spec_from_file_location(script.removesuffix(".py"), SCRIPTS / script)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


# ---------------------------------------------------------------------------------------------
# check_layers
# ---------------------------------------------------------------------------------------------


def test_layers_gate_fails_on_a_violation() -> None:
    result = _run("check_layers.py", "--root", str(FIXTURES / "layers_violation"))
    assert result.returncode == 1, result.stdout


@pytest.mark.parametrize(
    "expected_key",
    ["purity:domain", "direction:domain->adapters", "slice:lanzadera->expedientes"],
)
def test_layers_gate_names_each_violation_class(expected_key: str) -> None:
    result = _run("check_layers.py", "--root", str(FIXTURES / "layers_violation"))
    assert expected_key in result.stdout, result.stdout


def test_layers_gate_passes_on_clean_code() -> None:
    result = _run("check_layers.py", "--root", str(FIXTURES / "layers_clean"))
    assert result.returncode == 0, result.stdout


def test_layers_gate_fails_when_the_package_is_missing() -> None:
    """Fail closed: an empty walk must never report success."""
    result = _run("check_layers.py", "--root", str(FIXTURES))
    assert result.returncode == 1


# ---------------------------------------------------------------------------------------------
# check_complexity
# ---------------------------------------------------------------------------------------------


def test_complexity_gate_fails_above_the_ceiling() -> None:
    result = _run("check_complexity.py", "--root", str(FIXTURES / "complexity_violation"))
    assert result.returncode == 1, result.stdout


def test_complexity_gate_passes_below_the_ceiling() -> None:
    result = _run("check_complexity.py", "--root", str(FIXTURES / "complexity_clean"))
    assert result.returncode == 0, result.stdout


def test_complexity_ceiling_is_absolute_not_top_n() -> None:
    """Hard Rule 12: the verdict on a function must not depend on its neighbours.

    The violation fixture holds one offender next to many trivial functions. A ``top-N`` gate
    would rank the offender out of view; an absolute ceiling still reports it.
    """
    result = _run("check_complexity.py", "--root", str(FIXTURES / "complexity_violation"))
    assert "too_many_branches" in result.stdout, result.stdout


# ---------------------------------------------------------------------------------------------
# check_branch_name
# ---------------------------------------------------------------------------------------------


@pytest.mark.parametrize("branch", ["feat/42-lanzadera-auth", "fix/7-crlf-count", "main"])
def test_branch_name_gate_accepts_valid_names(branch: str) -> None:
    assert _run("check_branch_name.py", "--branch", branch).returncode == 0


@pytest.mark.parametrize(
    "branch",
    ["lanzadera-auth", "feat/lanzadera-auth", "feat/42_lanzadera_auth", "FEAT/42-auth", "develop"],
)
def test_branch_name_gate_rejects_invalid_names(branch: str) -> None:
    assert _run("check_branch_name.py", "--branch", branch).returncode == 1


# ---------------------------------------------------------------------------------------------
# check_pr_size
#
# The diff path needs a git repository, so it is covered by the workflow itself rather than here.
# The override policy is pure logic and is covered directly: an override with no stated reason
# must not count as an override.
# ---------------------------------------------------------------------------------------------


def test_pr_size_override_requires_a_reason() -> None:
    module = _load("check_pr_size.py")
    assert module.override_reason("size:exception") is None
    assert module.override_reason("") is None
    assert (
        module.override_reason("size:exception\nsize-exception-reason: generated migration")
        == "generated migration"
    )


# ---------------------------------------------------------------------------------------------
# check_crap
# ---------------------------------------------------------------------------------------------


def test_crap_gate_fails_on_untested_complexity() -> None:
    result = _run("check_crap.py", "--root", str(FIXTURES / "crap_violation"))
    assert result.returncode == 1, result.stdout


def test_crap_gate_catches_what_the_complexity_ceiling_lets_through() -> None:
    """The whole argument for CRAP, pinned as a test.

    The offender sits at complexity 4 — far under the complexity ceiling of 15, so that gate
    passes it. With no tests behind it, CRAP scores it 20. If these two ever agree, the CRAP
    gate has stopped adding information.
    """
    complexity = _run("check_complexity.py", "--root", str(FIXTURES / "crap_violation"))
    crap = _run("check_crap.py", "--root", str(FIXTURES / "crap_violation"))
    assert complexity.returncode == 0, complexity.stdout
    assert crap.returncode == 1, crap.stdout
    assert "CRAP 20.0" in crap.stdout


def test_crap_gate_passes_on_small_covered_functions() -> None:
    result = _run("check_crap.py", "--root", str(FIXTURES / "crap_clean"))
    assert result.returncode == 0, result.stdout


def test_crap_gate_fails_closed_without_coverage_data(tmp_path) -> None:
    """No coverage data means no verdict, and no verdict must never read as success."""
    package = tmp_path / "app" / "lanzadera" / "domain"
    package.mkdir(parents=True)
    (package / "model.py").write_text("def f():\n    return 1\n", encoding="utf-8")
    result = _run("check_crap.py", "--root", str(tmp_path))
    assert result.returncode == 1
    assert "coverage" in (result.stderr + result.stdout).lower()


# ---------------------------------------------------------------------------------------------
# check_dry
# ---------------------------------------------------------------------------------------------


def test_dry_gate_fails_on_a_renamed_copy_paste() -> None:
    result = _run("check_dry.py", "--root", str(FIXTURES / "dry_violation"))
    assert result.returncode == 1, result.stdout
    assert "alpha.py" in result.stdout and "beta.py" in result.stdout


def test_dry_gate_passes_on_distinct_code() -> None:
    result = _run("check_dry.py", "--root", str(FIXTURES / "dry_clean"))
    assert result.returncode == 0, result.stdout


def test_dry_detection_is_stable_across_processes() -> None:
    """Determinism pin: the digest must not depend on PYTHONHASHSEED.

    An earlier draft keyed clone groups on the builtin ``hash()``, which is salted per process.
    The gate then reported different groups on identical code from one run to the next.
    """
    first = _run("check_dry.py", "--root", str(FIXTURES / "dry_violation"), "--json")
    second = _run("check_dry.py", "--root", str(FIXTURES / "dry_violation"), "--json")
    assert first.stdout == second.stdout


# ---------------------------------------------------------------------------------------------
# quality_report — the indicator layer
# ---------------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "gate_script",
    [
        "check_layers.py",
        "check_complexity.py",
        "check_crap.py",
        "check_dry.py",
        "check_branch_name.py",
    ],
)
def test_every_gate_emits_a_well_formed_envelope(gate_script: str) -> None:
    root = FIXTURES / "crap_clean"
    args = ["--branch", "main"] if gate_script == "check_branch_name.py" else ["--root", str(root)]
    result = _run(gate_script, *args, "--json")
    envelope = json.loads(result.stdout)
    assert envelope["gate"]
    assert envelope["status"] in {"pass", "fail", "error"}
    assert isinstance(envelope["indicators"], dict)
    assert isinstance(envelope["findings"], list)


def test_quality_report_aggregates_indicators(tmp_path) -> None:
    out = tmp_path / "quality-report.json"
    result = _run(
        "quality_report.py", "--root", str(FIXTURES / "crap_clean"), "--out", str(out)
    )
    assert result.returncode == 0, result.stdout
    report = json.loads(out.read_text(encoding="utf-8"))
    assert report["status"] == "pass"
    for indicator in ("layers.violations", "complexity.max_complexity", "crap.max_crap"):
        assert indicator in report["indicators"], report["indicators"].keys()
    assert report["indicators"]["crap.max_crap"]["ceiling"] == 6.0


def test_quality_report_names_the_failing_gate(tmp_path) -> None:
    out = tmp_path / "quality-report.json"
    result = _run(
        "quality_report.py", "--root", str(FIXTURES / "crap_violation"), "--out", str(out)
    )
    assert result.returncode == 1
    report = json.loads(out.read_text(encoding="utf-8"))
    assert report["failed_gates"] == ["crap"]


def test_quality_report_is_byte_identical_for_the_same_commit(tmp_path) -> None:
    """Determinism pin: no wall-clock timestamp may leak into the report."""
    first, second = tmp_path / "a.json", tmp_path / "b.json"
    _run("quality_report.py", "--root", str(FIXTURES / "crap_clean"), "--out", str(first))
    _run("quality_report.py", "--root", str(FIXTURES / "crap_clean"), "--out", str(second))
    assert first.read_bytes() == second.read_bytes()
