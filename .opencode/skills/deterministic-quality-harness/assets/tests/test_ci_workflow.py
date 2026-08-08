# HARNESS-PROVENANCE: deterministic-quality-harness v1.2 — assets/tests/test_ci_workflow.py
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
REQUIRED_COMMANDS = (
    "ruff format --check",
    "ruff check .",
    "mypy app/",
    "pytest --cov --cov-report=json:coverage.json",
    "python scripts/quality_report.py",
    "python scripts/check_pr_size.py",
    "python scripts/check_branch_name.py",
    "pip-audit",
    "gitleaks",
    "trivy",
)

#: The four code gates, in the order they must run. Deduplication moves code, which changes
#: complexity, which changes CRAP — so the sequence is part of the contract, not a preference.
REQUIRED_GATE_ORDER = ("layers", "complexity", "crap", "dry")

_SHA_PIN = re.compile(r"^[0-9a-f]{40}$")


def _find_workflow() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / ".github" / "workflows" / "ci.yml"
        if candidate.is_file():
            return candidate
    # Fallback: the skill's own asset layout, so the template is self-testing.
    candidate = here.parent.parent / "ci.yml"
    if candidate.is_file():
        return candidate
    raise AssertionError("ci.yml not found")


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
    """Hard Rule 13: the order lives in quality_report.GATES, and this is what pins it.

    Keeping the four code gates inside one script rather than four YAML steps is deliberate: a
    reviewer can reorder two YAML steps without noticing, and the resulting verdict would depend
    on the order in which someone happened to fix things.
    """
    scripts = Path(__file__).resolve().parent.parent / "scripts" / "quality_report.py"
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
