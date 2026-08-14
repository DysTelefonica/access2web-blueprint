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
    "ruff format --check --config app/pyproject.toml",
    "ruff check --config app/pyproject.toml .",
    "mypy --explicit-package-bases app/",
    "python scripts/check_workflows.py",
    "pytest -c app/pyproject.toml --rootdir=app --cov --cov-report=json:coverage.json",
    "python scripts/check_workflows.py",
    "python scripts/quality_report.py",
    "python scripts/check_mutation.py",
    "python scripts/check_pr_size.py",
    "python scripts/check_branch_name.py",
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

#: The code gates, in the order they must run. Deduplication moves code, which changes
#: complexity — so the sequence is part of the contract, not a preference.
REQUIRED_GATE_ORDER = (
    "layers",
    "complexity",
    "mutation_sites",
    "dry",
    # DA-13, last: a symbol walker, not a metric, so it neither consumes nor
    # invalidates the numbers before it.
    "legacy_hashes",
)

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
    """Hard Rule 15: `ubuntu-latest` is a different machine from one month to the next.

    Each label is checked on its own. `runs-on` is a scalar for a hosted runner and a
    LIST for a self-hosted one, and the previous version stringified the whole value
    and asked whether it ended in `-latest`. A list ends in `]`, so from the moment
    this repository moved to self-hosted labels (#113) the check passed unconditionally
    — including for `[self-hosted, ubuntu-latest]`. A gate that cannot fail is not a
    gate, and this one stopped being able to fail the day it mattered most.
    """
    floating = []
    for name, job in workflow["jobs"].items():
        runs_on = job.get("runs-on", "")
        labels = runs_on if isinstance(runs_on, list) else [runs_on]
        for label in labels:
            if str(label).endswith("-latest"):
                floating.append(f"{name}: {label}")
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
        if not stripped or stripped.startswith(("#", ".")) or ":" not in stripped or " " in head:
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
    expected = [command for command in REQUIRED_COMMANDS if command not in VERIFY_EXCLUSIONS]
    # `ruff check .` reaches the recipe as `python -m ruff check .`; compare on
    # the invariant tail so the Makefile stays free to route through $(RUFF).
    missing = [
        command
        for command in expected
        if command.removeprefix("python ") not in verify_commands.replace("python -m ", "")
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

    Slice 4 of issue #266 (architectural-guards-over-metrics) lands the
    ``check_decision_guards.py`` script ahead of its CI wiring. The gate is
    invoked standalone (unit tests on synthetic markdown) and waits for slice 6
    to add it to ``GATES`` in ``quality_report.py``. Until then
    ``KNOWN_PENDING_GATES`` is the documented exception that keeps this guard
    honest about the slice boundary without forcing slice 4 to do slice 6 work.
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

    unwired = sorted(path.name for path in scripts_dir.glob("check_*.py") if path.name not in wired)
    assert not unwired, (
        f"these gate scripts exist but no CI step runs them: {unwired}. A gate that "
        "runs nowhere is a false guarantee — wire it into ci.yml or delete it "
        "(Decision Gate: 'Gate runs locally but not in CI'). Pending slice "
        "exceptions live in KNOWN_PENDING_GATES."
    )


#: Gate scripts that exist on disk but have not been wired into CI yet because
#: they are intentionally staged ahead of their aggregator entry. Each entry
#: MUST name the slice that removes the exception, otherwise the exception
#: becomes a permanent dead-script tomb.
KNOWN_PENDING_GATES: tuple[str, ...] = (
    # Slice 4 of #266: scripts/check_decision_guards.py is the unit-tested gate.
    # Slice 6 wires it into GATES and removes this entry.
    "check_decision_guards.py",
)


def _install_blocks(run_blocks: list[str]) -> list[str]:
    """Run blocks that target the platform project for `pip install`.

    An install block is any run block that combines `pip install` with the
    `app[dev]` editable target — this is what every quality, mutation, and
    security job calls when wiring the project into the runner's venv.
    """
    return [block for block in run_blocks if "pip install" in block and "app[dev]" in block]


def test_install_step_defends_against_missing_pyproject(run_blocks: list[str]) -> None:
    """F2 of #117: a chain that ships pyproject.fragment.toml without
    app/pyproject.toml must NOT fail with a cryptic pip error.

    Every install step targeting `app[dev]` MUST guard the install with a
    file existence check, harden the shell with `set -euo pipefail`, exit
    non-zero on the failure path, and reference issue #117 in the diagnostic
    so an operator staring at a red job can recognise the chain-hole shape.
    """
    installs = _install_blocks(run_blocks)
    assert installs, "no run block installs app[dev] — this test is stale"

    for i, block in enumerate(installs, start=1):
        # Defensive guard: prove the project exists before pip sees it.
        assert "app/pyproject.toml" in block, (
            f"install step #{i} does not check for app/pyproject.toml: {block}"
        )
        # Hardening: without `set -euo pipefail` the guard's `exit 1` can be
        # masked by a subsequent command, and the install step stops being
        # able to fail (Hard Rule 1).
        assert "set -euo pipefail" in block, (
            f"install step #{i} does not harden its shell with set -euo pipefail: {block}"
        )
        # Hard Rule 1: no silent pass on a real failure.
        assert "|| true" not in block, f"install step #{i} uses '|| true': {block}"
        # The failure path must terminate the job, not let it limp along.
        assert "exit 1" in block, f"install step #{i} does not exit non-zero on failure: {block}"
        # Diagnostic must point operators to the umbrella issue. The exact
        # spelling varies; the substring `#117` is the canonical anchor.
        assert "#117" in block, (
            f"install step #{i} does not reference issue #117 in its failure path: {block}"
        )


def test_install_step_reports_chain_hole_when_fragment_only(
    run_blocks: list[str],
) -> None:
    """The fragment-only branch must produce a diagnostic naming F2.

    When `app/pyproject.toml` is missing but `pyproject.fragment.toml` is
    present at the repo root, the install step MUST name that file in the
    failure message. Otherwise an operator has no way to recognise the
    chain-hole shape and link the failure back to the umbrella issue.
    """
    installs = _install_blocks(run_blocks)
    assert installs, "no run block installs app[dev] — this test is stale"
    for i, block in enumerate(installs, start=1):
        assert "pyproject.fragment.toml" in block, (
            f"install step #{i} does not mention pyproject.fragment.toml: {block}"
        )


def test_install_step_python_pin_unchanged(workflow: dict) -> None:
    """Hard Rule 15: the interpreter is part of the gate.

    Any patch to the Python pin is a different verdict on unchanged code, so
    a fix for the install step must not drift the pin (Hard Rule 15).
    """
    assert workflow["env"]["PYTHON_VERSION"] == "3.12.11", (
        f"PYTHON_VERSION drifted from 3.12.11 to {workflow['env']['PYTHON_VERSION']!r}"
    )


def test_secret_scan_proves_it_scanned_something() -> None:
    """Hard Rule 18: a zero-byte scan is not a clean tree.

    gitleaks went green after inspecting 0 bytes when the workspace was not visible
    to the host Docker daemon. trivy, given the identical broken mount, failed loudly
    — because it looks for one named file, while gitleaks walks a tree and an empty
    tree has no secrets in it.

    That difference is not luck. Any gate that inspects a SET treats the empty set as
    success unless someone teaches it otherwise, and a secret scanner is the worst
    possible place to learn that lesson late.
    """
    workflows = _find_workflow().parent
    security = (workflows / "security.yml").read_text(encoding="utf-8")
    assert "GITLEAKS_MIN_BYTES" in security, (
        "the gitleaks job must check how many bytes it inspected before accepting its "
        "verdict (issue #115, Hard Rule 18)"
    )
    assert "scanned ~" in security, (
        "the liveness proof must read gitleaks' own reported scan volume, not infer it"
    )

    deep = (workflows / "security-deep.yml").read_text(encoding="utf-8")
    assert "rev-list --count" in deep, (
        "the full-history scan must prove it had history to walk; a scan over zero "
        "commits reports clean as convincingly as one over ten thousand"
    )
