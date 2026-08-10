# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — tests/test_ci_workflow.py
"""Wiring pin for the CI workflow (Hard Rule 4).

A gate that exists only as a script is a gate nobody runs. These tests parse the workflow and
fail when a step is deleted, when a gate is neutered, or when a pin goes floating. They are the
reason a silent revert cannot pass review.
"""

from __future__ import annotations

import importlib.util
import re
from pathlib import Path

import pytest
import yaml

#: Every gate that must be wired, identified by the command it runs.
#:
#: ``test_no_legacy_compat.py`` (Phase 6) and `check_mutation.py` (Phase 2+) are NOT in this
#: list — Phase 0 only ships the gates whose scripts are present in `scripts/`.
REQUIRED_COMMANDS = (
    "ruff format --check",
    "ruff check .",
    "mypy src/",
    "pytest --cov=app",
    "python scripts/quality_report.py",
    "python scripts/check_pr_size.py",
    "python scripts/check_branch_name.py",
    "pip-audit",
)

#: The four code gates, in the order they must run. Deduplication moves code, which changes
#: complexity, which changes CRAP — so the sequence is part of the contract, not a preference.
#: `legacy_hashes` is a fifth gate (DA-13) that fires as part of `quality_report.py` and
#: stays alongside the others in the fixed execution order.
REQUIRED_GATE_ORDER = ("layers", "complexity", "crap", "dry", "legacy_hashes")

_SHA_PIN = re.compile(r"^[0-9a-f]{40}$")


def _find_workflow() -> Path:
    here = Path(__file__).resolve()
    # The test lives at `<root>/tests/test_ci_workflow.py`; the workflow lives at
    # `<root>/.github/workflows/ci.yml`.
    for parent in here.parents:
        candidate = parent / ".github" / "workflows" / "ci.yml"
        if candidate.is_file():
            return candidate
    raise AssertionError("ci.yml not found")


def _find_quality_report() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / "scripts" / "quality_report.py"
        if candidate.is_file():
            return candidate
    raise AssertionError("scripts/quality_report.py not found")


@pytest.fixture(scope="module")
def workflow() -> dict:
    return yaml.safe_load(_find_workflow().read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def steps(workflow: dict) -> list[dict]:
    collected: list[dict] = []
    for job in workflow["jobs"].values():
        collected.extend(job.get("steps", []))
    return collected


@pytest.fixture(scope="module")
def run_blocks(steps: list[dict]) -> list[str]:
    return [step["run"] for step in steps if "run" in step]


@pytest.mark.parametrize("command", REQUIRED_COMMANDS)
def test_gate_is_wired(run_blocks: list[str], command: str) -> None:
    assert any(command in block for block in run_blocks), f"no CI step runs '{command}'"


def test_no_step_swallows_its_exit_code(steps: list[dict]) -> None:
    """Hard Rule 1: a gate that cannot fail is not a gate."""
    offenders = [step.get("name", "<unnamed>") for step in steps if step.get("continue-on-error")]
    assert not offenders, f"continue-on-error found on: {offenders}"


def test_no_run_block_forces_success(run_blocks: list[str]) -> None:
    """Hard Rule 1, the other half: `|| true` neuters a scanner while it still looks green."""
    offenders = [block for block in run_blocks if "|| true" in block]
    assert not offenders, f"'|| true' found in: {offenders}"


def test_every_action_is_pinned_to_a_sha(steps: list[dict]) -> None:
    """Hard Rule 15: a floating tag turns a green gate red with no code change."""
    unpinned = []
    for step in steps:
        uses = step.get("uses")
        if not uses:
            continue
        _, _, ref = uses.partition("@")
        if not _SHA_PIN.match(ref):
            unpinned.append(uses)
    assert not unpinned, f"actions not pinned to a 40-character SHA: {unpinned}"


def test_gate_order_is_pinned_in_code() -> None:
    """Hard Rule 13: the order lives in quality_report.GATES, and this is what pins it."""
    scripts = _find_quality_report()
    spec = importlib.util.spec_from_file_location("quality_report", scripts)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    assert tuple(name for name, _, _ in module.GATES) == REQUIRED_GATE_ORDER


def test_no_job_runs_on_a_floating_runner(workflow: dict) -> None:
    """Hard Rule 15: `ubuntu-latest` is a different machine from one month to the next."""
    floating = [
        f"{name}: {job['runs-on']}"
        for name, job in workflow["jobs"].items()
        if str(job.get("runs-on", "")).endswith("-latest")
    ]
    assert not floating, f"jobs on a floating runner label: {floating}"


def test_security_workflow_pins_digests() -> None:
    """Hard Rule 6: scanner images use sha256:..., not tags."""
    here = Path(__file__).resolve()
    security = here.parent.parent / ".github" / "workflows" / "security.yml"
    deep = here.parent.parent / ".github" / "workflows" / "security-deep.yml"
    for path in (security, deep):
        assert path.is_file(), f"missing workflow: {path.name}"
        contents = path.read_text(encoding="utf-8")
        # Concatenate continuation lines (`run: >`) so a multi-line `docker run`
        # block is treated as one entry. The pin must live somewhere in the
        # block; YAML folds make per-line scanning unreliable.
        folded: list[str] = []
        for line in contents.splitlines():
            if line.startswith("          ") or line.startswith("        "):
                if folded:
                    folded[-1] += " " + line.strip()
                else:
                    folded.append(line.strip())
            elif line.strip():
                folded.append(line.strip())
        for entry in folded:
            if "docker run" not in entry:
                continue
            assert "@sha256:" in entry, (
                f"{path.name} uses a non-digest image reference: {entry[:200]}"
            )