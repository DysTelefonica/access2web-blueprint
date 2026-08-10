# HARNESS-PROVENANCE: deterministic-quality-harness v1.5 — assets/tests/test_ci_workflow.py
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
    "python scripts/check_mutation.py",
    "python scripts/check_pr_size.py",
    "python scripts/check_branch_name.py",
    "pip-audit",
    "gitleaks",
    "trivy",
)

#: Gates CI runs that `make verify` deliberately does not (Hard Rule 19).
#: Each needs something a developer workstation does not have: the pull-request
#: payload, a weekly time budget, or a Docker daemon. Listing them here is what
#: turns an absence into a recorded decision — anything NOT in this tuple must
#: be reachable from `make verify`.
VERIFY_EXCLUSIONS = (
    "python scripts/check_pr_size.py",  # needs merge-base + labels
    "python scripts/check_branch_name.py",  # needs the head ref
    "python scripts/check_mutation.py",  # weekly schedule, far too slow per PR
    "pip-audit",  # network: advisory database
    "gitleaks",  # docker
    "trivy",  # docker
)

#: The four code gates, in the order they must run. Deduplication moves code, which changes
#: complexity, which changes CRAP — so the sequence is part of the contract, not a preference.
REQUIRED_GATE_ORDER = ("layers", "complexity", "crap", "mutation_sites", "dry")

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
    offenders = [
        step.get("name", "<unnamed>") for step in steps if step.get("continue-on-error")
    ]
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


# --- Set-level parity: scripts <-> ci.yml <-> make verify (Hard Rule 19) ---
#
# Every test above pins ONE gate to ONE step. That catches a deletion and
# nothing else: a gate script written and never wired passes, because the
# hardcoded entry for it was never added either. The three tests below pin the
# SETS, which is the only assertion that survives someone adding gate number
# eighteen.


def _find_makefile() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / "Makefile"
        if candidate.is_file():
            return candidate
    # Fallback: the skill's own asset layout, so the template is self-testing.
    candidate = here.parent.parent / "Makefile"
    if candidate.is_file():
        return candidate
    raise AssertionError("Makefile not found")


def _parse_makefile(path: Path) -> dict[str, tuple[list[str], list[str]]]:
    """Parse a Makefile into ``{target: (prerequisites, recipe lines)}``.

    Deliberately minimal — enough to walk the ``verify`` dependency graph.
    Backslash continuations are joined first so a wrapped prerequisite list
    reads as one logical line; comments, ``.PHONY`` and variable assignments
    are skipped.
    """
    logical: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if logical and logical[-1].endswith("\\"):
            logical[-1] = logical[-1].removesuffix("\\").rstrip() + " " + line.strip()
        else:
            logical.append(line)

    targets: dict[str, tuple[list[str], list[str]]] = {}
    current: str | None = None
    for line in logical:
        if line.startswith("\t"):
            if current is not None:
                targets[current][1].append(line.strip())
            continue
        stripped = line.strip()
        head = stripped.partition(":")[0]
        if (
            not stripped
            or stripped.startswith(("#", "."))
            or ":" not in stripped
            or " " in head
        ):
            current = None
            continue
        current = head.strip()
        targets[current] = (stripped.partition(":")[2].split(), [])
    return targets


@pytest.fixture(scope="module")
def verify_commands() -> str:
    """Every command reachable from ``make verify``, with variables expanded."""
    path = _find_makefile()
    targets = _parse_makefile(path)
    assert "verify" in targets, f"{path} defines no `verify` target (Hard Rule 19)"

    seen: set[str] = set()
    collected: list[str] = []

    def walk(name: str) -> None:
        if name in seen or name not in targets:
            return
        seen.add(name)
        prerequisites, recipe = targets[name]
        for prerequisite in prerequisites:
            walk(prerequisite)
        collected.extend(recipe)

    walk("verify")
    blob = "\n".join(collected)
    for variable, expansion in (
        ("$(PYTHON)", "python"),
        ("$(RUFF)", "python -m ruff"),
        ("$(MYPY)", "python -m mypy"),
        ("$(PYTEST)", "python -m pytest"),
        ("$(PACKAGE)", "app"),
    ):
        blob = blob.replace(variable, expansion)
    return blob


def test_make_verify_runs_every_local_gate(verify_commands: str) -> None:
    """Hard Rule 19: one command must equal green.

    Without this, the gate list lives only in ci.yml, `make verify` silently
    becomes a subset of it, and a green local run stops meaning anything. The
    subset always drifts downward, because nothing measures it.

    Anything CI runs that a workstation genuinely cannot must be recorded in
    VERIFY_EXCLUSIONS — an absence that is written down is a decision, and an
    absence that is not is a hole.
    """
    expected = [
        command for command in REQUIRED_COMMANDS if command not in VERIFY_EXCLUSIONS
    ]
    # `ruff check .` reaches the recipe as `python -m ruff check .`; compare on
    # the invariant tail so the Makefile stays free to route through $(RUFF).
    missing = [
        command
        for command in expected
        if command.removeprefix("python ")
        not in verify_commands.replace("python -m ", "")
        and command not in verify_commands
    ]
    assert not missing, (
        f"ci.yml gates a pull request on {missing}, but `make verify` never runs them. "
        "Add a target per gate and list it in the `verify` prerequisites, in the same "
        "order ci.yml runs it (Hard Rule 19)."
    )


def test_make_verify_excludes_what_a_workstation_cannot_run(
    verify_commands: str,
) -> None:
    """The exclusions are a design decision, so they are pinned like any other.

    Folding the weekly mutation session or the Docker scanners into ``verify``
    would make the local gate slow, Linux-only, or impossible without a daemon
    — and a gate people stop running is worse than one that never existed. If
    one of them ever becomes cheap enough to include, deleting its entry from
    VERIFY_EXCLUSIONS is the deliberate act that records the change.
    """
    leaked = [command for command in VERIFY_EXCLUSIONS if command in verify_commands]
    assert not leaked, (
        f"`make verify` runs {leaked}, which VERIFY_EXCLUSIONS declares out of scope. "
        "Either remove it from verify or remove it from the exclusion list — silently "
        "disagreeing with the declared contract is the drift Hard Rule 19 exists to stop."
    )


def test_every_gate_script_on_disk_is_wired_in_ci(run_blocks: list[str]) -> None:
    """Hard Rule 4, at the level of the set rather than one gate at a time.

    REQUIRED_COMMANDS catches a gate being deleted from ci.yml. It cannot catch
    a gate that was written and never wired, because whoever forgot the CI step
    also forgot to add the entry here. Walking `scripts/` closes that: the gate
    exists on disk, so it must run somewhere, or be deleted.
    """
    scripts_dir = _find_makefile().parent / "scripts"
    if not scripts_dir.is_dir():
        pytest.skip("no scripts/ directory in this layout")

    # A gate counts as wired either by its own CI step or through the
    # aggregator: the four code gates run inside quality_report.py on purpose
    # (Hard Rule 13 — their order is load-bearing and belongs in code).
    wired = "\n".join(run_blocks)
    aggregator = scripts_dir / "quality_report.py"
    if aggregator.is_file():
        wired += "\n" + aggregator.read_text(encoding="utf-8")

    unwired = sorted(
        path.name for path in scripts_dir.glob("check_*.py") if path.name not in wired
    )
    assert not unwired, (
        f"these gate scripts exist but no CI step runs them: {unwired}. A gate that "
        "runs nowhere is a false guarantee — wire it into ci.yml or delete it "
        "(Decision Gate: 'Gate runs locally but not in CI')."
    )
