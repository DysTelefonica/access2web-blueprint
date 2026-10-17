# HARNESS-PROVENANCE: deterministic-quality-harness v1.5 — assets/tests/test_ci_workflow.py
"""Wiring pin for the CI workflow (Hard Rule 4).

A gate that exists only as a script is a gate nobody runs. These tests parse the workflow and
fail when a step is deleted, when a gate is neutered, or when a pin goes floating. They are the
reason a silent revert cannot pass review.
"""

from __future__ import annotations

import importlib.util
import re
import tomllib
from pathlib import Path

import pytest
import yaml

#: Every gate that must be wired, identified by the command it runs.
REQUIRED_COMMANDS = (
    "ruff format --check --config app/pyproject.toml",
    "ruff check --config app/pyproject.toml .",
    "mypy --explicit-package-bases app/",
    "python scripts/check_workflows.py",
    'pytest -c app/pyproject.toml --rootdir=app -m "not integration"',
    "python scripts/check_workflows.py",
    "scripts/quality_report.py",
    "python scripts/check_pr_size.py",
    "python scripts/check_walkthrough_schema.py --strict",
    "python scripts/check_required_jobs.py",
)

#: Gates CI runs that `make verify` deliberately does not (Hard Rule 19).
#: Each needs something a developer workstation does not have: the pull-request
#: payload, a weekly time budget, or a Docker daemon. Listing them here is what
#: turns an absence into a recorded decision — anything NOT in this tuple must
#: be reachable from `make verify`.
VERIFY_EXCLUSIONS = (
    "mypy --explicit-package-bases app/",  # full suite: CI quality job only
    "python scripts/check_workflows.py",  # ci.yml contract: CI quality job only
    'pytest -c app/pyproject.toml --rootdir=app -m "not integration"',  # full suite
    "scripts/quality_report.py",  # code gates after the suite: CI quality job only
    "pip-audit",  # network: advisory database
    "gitleaks",  # docker
    "trivy",  # docker
    # needs the real `needs:` context (toJSON(needs)) from a GitHub Actions
    # run; there is nothing meaningful to feed it on a workstation (#702).
    "python scripts/check_required_jobs.py",
)

#: The code gates, in the order they must run. Deduplication moves code, which changes
#: complexity — so the sequence is part of the contract, not a preference.
REQUIRED_GATE_ORDER = (
    "layers",
    "complexity",
    "mutation_sites",
    "dry",
    # DA-13, last among the metric gates: a symbol walker, not a metric,
    # so it neither consumes nor invalidates the numbers before it.
    "legacy_hashes",
    # L01 of #264 (expedientes-web-migration): retirement gate. Same
    # rationale as legacy_hashes (a structural walker, not a metric) so
    # it sits at the end of the chain.
    "legacy_retirement",
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


def test_dependabot_covers_python_and_github_actions() -> None:
    """Version updates must cover every dependency manifest in the repository."""
    dependabot = yaml.safe_load(
        (_find_workflow().parent.parent / "dependabot.yml").read_text(encoding="utf-8")
    )
    covered = {
        (update["package-ecosystem"], update["directory"]) for update in dependabot["updates"]
    }

    assert covered == {("pip", "/app"), ("github-actions", "/"), ("npm", "/e2e")}


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


def _local_preflight() -> str:
    """El texto del preflight local, donde vive la única lista de gates."""
    path = _find_makefile().parent / "scripts" / "local-preflight.sh"
    assert path.is_file(), f"{path} no existe: `make verify` delega en él (Hard Rule 19)"
    return path.read_text(encoding="utf-8")


def _local_preflight_commands() -> list[str]:
    """Los comandos del preflight, en orden y sin comentarios ni `set -`."""
    commands: list[str] = []
    for raw in _local_preflight().splitlines():
        line = raw.strip()
        if line and not line.startswith("#") and not line.startswith("set -"):
            commands.append(" ".join(line.split()))
    return commands


def test_make_verify_runs_every_local_gate(verify_commands: str) -> None:
    """Hard Rule 19: one command must equal green.

    Without this, the gate list lives only in ci.yml, `make verify` silently
    becomes a subset of it, and a green local run stops meaning anything. The
    subset always drifts downward, because nothing measures it.

    Anything CI runs that a workstation genuinely cannot must be recorded in
    VERIFY_EXCLUSIONS — an absence that is written down is a decision, and an
    absence that is not is a hole.
    """
    # 1. `make verify` no tiene su propia lista: delega en el script.
    assert "scripts/local-preflight.sh" in verify_commands, (
        "`make verify` does not delegate to scripts/local-preflight.sh: the canonical "
        "list of gates must live in exactly one place (Hard Rule 19)."
    )

    # 2. El script corre los mismos gates que CI, en el mismo orden.
    expected: list[str] = []
    for command in REQUIRED_COMMANDS:
        if command in VERIFY_EXCLUSIONS:
            continue
        if command.removeprefix("python ") in [
            v.removeprefix("python ") for v in VERIFY_EXCLUSIONS
        ]:
            continue
        if command not in expected:
            expected.append(command)

    script = " ".join(_local_preflight_commands())
    positions = [script.find(" ".join(command.split())) for command in expected]
    assert -1 not in positions, (
        "ci.yml gates a pull request on "
        f"{[c for c, pos in zip(expected, positions, strict=True) if pos == -1]}, but the local "
        "preflight never runs them. Add the gate to scripts/local-preflight.sh, in the "
        "same order ci.yml runs it (Hard Rule 19)."
    )
    assert positions == sorted(positions), (
        "the local preflight runs the CI gates in a different order: the sequence is "
        "part of the contract (complexity and coverage depend on it)."
    )

    # 3. Y no corre ningún gate que no esté declarado: cada línea del script es
    #    uno de los gates esperados.
    allowed = set()
    for command in expected:
        normalized = " ".join(command.split())
        allowed |= {normalized, normalized.removeprefix("python "), f"python {normalized}"}
    extra = [
        command
        for command in _local_preflight_commands()
        if not any(
            " ".join(command.split()) == candidate
            or " ".join(command.split()).startswith(candidate)
            for candidate in allowed
        )
    ]
    assert not extra, (
        f"the local preflight runs commands that are not in the CI gate contract: {extra}. "
        "Either wire them in ci.yml or record the decision in VERIFY_EXCLUSIONS."
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

    unwired = sorted(
        path.name
        for path in scripts_dir.glob("check_*.py")
        if path.name not in wired and path.name not in KNOWN_PENDING_GATES
    )
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
    # Pieza 4 of testing-strategy epic (#535): scripts/check_test_classification.py
    # ships in --dry-run mode for one sprint per testing-strategy.md §Pieza 4.3.
    # After the sprint, --strict activates in ci.yml's quality job and this entry
    # is removed. Issue #535 tracks the exit criteria.
    "check_test_classification.py",
    # Issue #857 (PR-B): the branch-name job leaves ci.yml, so the script has no
    # CI step. Issue #859 deletes scripts/check_branch_name.py and removes this entry.
    "check_branch_name.py",
    # Issue #857, PR-C1b: the mutation jobs leave ci.yml, so the scripts have no
    # CI step in the PR rollup. They stay on disk — the mutation job moved
    # to release.yml (PR-C1d) and reads them from there. The two
    # KNOWN_PENDING_GATES entries here mark the ci.yml-side absence.
    "check_mutation.py",
    "check_mutation_sites.py",
    # Issue #875: scripts/check_deadcode.py ships the ratchet without a CI step.
    # The quality-job wiring lives in the PR for the gentle-ai model (#857, owned
    # by apap4); that PR adds the step to ci.yml and removes this entry.
    "check_deadcode.py",
)


def _install_blocks(run_blocks: list[str]) -> list[str]:
    """Run blocks that install the pinned toolchain for `pip install`.

    An install block is any run block that combines `pip install` with the
    hash-pinned requirements file generated from `app/uv.lock` — this is what
    every quality and mutation job calls when wiring the project's toolchain
    into the runner's interpreter. The install no longer uses an editable
    `app[dev]` target: the project is imported through pytest's `pythonpath`
    and mypy's `--explicit-package-bases`, and the pattern's structural gate
    only verifies pins it can read (HR-26).
    """
    return [
        block
        for block in run_blocks
        if "pip install" in block and "app/requirements-dev.lock" in block
    ]


def test_every_setup_python_step_pins_the_same_literal(workflow: dict) -> None:
    """HR-26: el gate estructural solo puede leer un literal, no una expresión.

    Cada `actions/setup-python` lleva la versión exacta escrita en el paso —lo
    que el gate del patrón verifica— y ese literal tiene que ser el mismo que
    la variable declarada en `env`, para que el valor no derive en dos sitios.
    """
    declared = workflow["env"]["PYTHON_VERSION"]
    steps = [s for j in workflow["jobs"].values() for s in j.get("steps", [])]
    setup = [s for s in steps if str(s.get("uses", "")).startswith("actions/setup-python@")]

    assert setup, "no actions/setup-python step found"
    for step in setup:
        assert step["with"]["python-version"] == declared, (
            f"setup-python declares {step['with']['python-version']!r} and env declares {declared!r}"
        )


def test_install_step_defends_against_missing_pyproject(run_blocks: list[str]) -> None:
    """F2 of #117: a chain that ships pyproject.fragment.toml without
    app/pyproject.toml must NOT fail with a cryptic pip error.

    Every install step targeting the pinned toolchain MUST guard the install with a
    file existence check, harden the shell with `set -euo pipefail`, exit
    non-zero on the failure path, and reference issue #117 in the diagnostic
    so an operator staring at a red job can recognise the chain-hole shape.
    """
    installs = _install_blocks(run_blocks)
    assert installs, "no run block installs the pinned toolchain — this test is stale"

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
    assert installs, "no run block installs the pinned toolchain — this test is stale"
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


def test_release_e2e_generates_its_signing_secret_per_run() -> None:
    """The release smoke container must not reuse a plaintext signing key."""
    release = (_find_workflow().parent / "release.yml").read_text(encoding="utf-8")

    assert not re.search(r'-e SECRET_KEY="(?!\$)[^\"]+"', release)
    assert "secret=$(openssl rand -hex 32)" in release
    assert 'echo "::add-mask::$secret"' in release
    assert 'printf \'E2E_SECRET_KEY=%s\\n\' "$secret" >>"$GITHUB_ENV"' in release
    assert '-e SECRET_KEY="$E2E_SECRET_KEY"' in release


def test_required_job_replaces_merge_ready_and_aggregates_pr_reachable_jobs(
    workflow: dict,
) -> None:
    """Issue #702: `required` is the deterministic replacement for `merge-ready`.

    `needs:` cannot cross workflow files, so `security` and `codeql` are
    embedded as `uses: ./.github/workflows/*.yml` calls directly inside
    ci.yml — from GitHub's perspective those are still ordinary job ids the
    aggregator can depend on and read `needs.<id>.result` from. Issue #857,
    PR-D: the event set that feeds the aggregator drops `merge_group` (the
    queue is gone with the gentle-ai model) and gains `push` (the post-merge
    safety net on `main`).
    """
    jobs = workflow["jobs"]
    assert "merge-ready" not in jobs, "merge-ready was fully superseded by `required`"

    required = jobs["required"]
    # El `if` conserva la semántica de `always()` —un dependency saltado no
    # salta este check, que es el requerido de la protección de rama— con la
    # forma que el gate del patrón puede probar (HR-30): comparaciones de
    # `event_name` y `!cancelled()`, sin `always()` ni `!=` sueltos.
    assert required["if"] == (
        "(github.event_name == 'pull_request' || github.event_name == 'push' "
        "|| github.event_name == 'schedule' || github.event_name == 'workflow_dispatch') "
        "&& !cancelled()"
    )
    assert required["needs"] == ["pr-gate", "quality", "integration", "security", "codeql"]

    assert jobs["security"]["uses"] == "./.github/workflows/security.yml"
    assert jobs["codeql"]["uses"] == "./.github/workflows/codeql.yml"
    assert jobs["codeql"]["permissions"]["security-events"] == "write"

    step = required["steps"][-1]
    assert step["run"] == "python scripts/check_required_jobs.py"
    assert step["env"]["CI_NEEDS_JSON"] == "${{ toJSON(needs) }}"
    assert step["env"]["CI_EVENT_NAME"] == "${{ github.event_name }}"


def test_pr_gate_is_the_cheap_front_and_review_budget_is_gone(workflow: dict) -> None:
    """Issue #857, PR-B: the cheap front gate is `pr-gate`; `review-budget` is renamed away."""
    jobs = workflow["jobs"]
    assert "pr-gate" in jobs
    assert "review-budget" not in jobs


def test_integration_job_runs_pytest_m_integration_against_postgres_service() -> None:
    """Issue #882, PR-3: the integration job wires ci.yml to the in-job Postgres service.

    The DSN fixture (PR-2) fails closed when the integration URL is missing;
    this job provides that URL by running postgres as a service so the
    integration tests have a real database to talk to. The contract is
    three things at once:

      1. The job exists and declares ``services: postgres`` (so a
         later rename to an external service changes the contract).
      2. The image is pinned by SHA256 digest (the image runs as the
         pin, not whatever tag a publisher swaps under it later — ``pull_policy:
         always`` makes the digest load-bearing, a stale local cache would
         be a silent replay vulnerability in a release gate).
      3. The pytest command runs the full integration suite with no path
         filter — ``pytest -m integration`` — because partial coverage
         is what made this gate a false guarantee in the first place.
    """
    workflow = yaml.safe_load(_find_workflow().read_text(encoding="utf-8"))
    job = workflow["jobs"]["integration"]

    # (1) services: postgres exists. The exact service name is part of
    # the contract: it is referenced from the body of the step that
    # resolves the DSN from the assigned host port.
    assert "services" in job, (
        "`integration` job must declare `services:` (issue #882, PR-3): "
        "without a Postgres service the integration tests have no DSN"
    )
    assert "postgres" in job["services"], (
        "`integration.services.postgres` is the contract (issue #882, PR-3). "
        "Got: " + repr(list(job["services"].keys()))
    )
    postgres = job["services"]["postgres"]

    # (2) image is pinned by SHA256 digest. A bare tag is exactly the
    # supply-chain drift this guard exists to detect; a substring match
    # is fine because the digest format is fixed-length and unambiguous.
    image = postgres["image"]
    assert "@sha256:" in image, (
        f"`integration.services.postgres.image` must be pinned by SHA256 "
        f"digest (issue #882, PR-3). Got: {image!r}. A bare tag like "
        f"`postgres:16-alpine` is third-party code that can change between "
        f"two otherwise identical runs."
    )

    # (3) the pytest run is `-m integration` with no path filter. A path
    # filter on the integration job would re-introduce the partial-coverage
    # hole that the levels comment on #882 explicitly rejects.
    pytest_steps = [step for step in job["steps"] if "pytest" in step.get("run", "")]
    assert pytest_steps, (
        "`integration` job must contain at least one `pytest` step "
        "(issue #882, PR-3). Without one, the job is a green rollup that "
        "never asserts anything."
    )
    runs = [step["run"] for step in pytest_steps]
    assert any("-m integration" in run for run in runs), (
        f"`integration` pytest step must select `integration` marker "
        f"(issue #882, PR-3). Got runs: {runs!r}. No-path-filter is the "
        f"point: a subset of the suite is what made this gate a false "
        f"guarantee in the first place."
    )


def test_integration_job_runs_alembic_upgrade_before_pytest() -> None:
    """Issue #882, PR-3: the integration suite expects an up-to-date schema.

    The test files (``test_migration_0001.py``, ``test_user_repository_pg.py``)
    assume a clean schema. An alembic upgrade step BEFORE pytest is the
    contract; without it the integration suite is testing against a stale
    schema and the migration itself is untested.
    """
    workflow = yaml.safe_load(_find_workflow().read_text(encoding="utf-8"))
    job = workflow["jobs"]["integration"]
    step_names = [step.get("name", "") for step in job["steps"]]

    # The "Apply migrations" step name is the one the integration suite
    # expects; if a rename moves the contract here the test catches it.
    migration_steps = [step for step in job["steps"] if "alembic" in step.get("run", "")]
    assert migration_steps, (
        "`integration` job must run `alembic upgrade` before pytest "
        "(issue #882, PR-3). Without it, the suite tests against a "
        "stale schema and the migration itself is untested in CI."
    )
    pytest_steps = [step for step in job["steps"] if "pytest" in step.get("run", "")]
    assert pytest_steps, "pytest step expected (covered by another test)"

    # The migration step must appear BEFORE the pytest step in the
    # job's steps list. ``steps`` preserves YAML order.
    migration_index = job["steps"].index(migration_steps[0])
    pytest_index = job["steps"].index(pytest_steps[0])
    assert migration_index < pytest_index, (
        f"`alembic upgrade` (index {migration_index}) must run before "
        f"`pytest` (index {pytest_index}) in the integration job "
        f"(issue #882, PR-3). Step names: {step_names!r}"
    )


def test_heavy_jobs_wait_for_pr_gate(workflow: dict) -> None:
    """Issue #857, PR-B: a failed gate must not pay for quality, security, codeql or integration."""
    jobs = workflow["jobs"]
    for job in ("quality", "security", "codeql", "integration"):
        assert jobs[job]["needs"] == "pr-gate", f"`{job}` must declare `needs: pr-gate`"


def test_heavy_jobs_run_on_every_event_after_pr_gate_passed_or_skipped(
    workflow: dict,
) -> None:
    """Issue #857, PR-D: the heavy jobs run on every event the workflow listens to.

    `pr-gate` only fires on `pull_request`, so on `push` (post-merge safety
    net on `main`), `schedule` (nightly) and `workflow_dispatch` (manual)
    it is `skipped`. The heavy jobs that `needs: pr-gate` must accept the
    skip on every one of those events. `merge_group` is gone: the queue
    left with the gentle-ai model. Issue #882, PR-3 extends the same
    shape to `integration` so the in-job Postgres service runs after the
    merge and on the nightly as well — not just on the PR event.
    """
    jobs = workflow["jobs"]
    expected = (
        "(github.event_name == 'pull_request' || github.event_name == 'push' "
        "|| github.event_name == 'schedule' || github.event_name == 'workflow_dispatch') "
        "&& !cancelled() "
        "&& (needs.pr-gate.result == 'success' || needs.pr-gate.result == 'skipped')"
    )
    for job in ("quality", "security", "codeql", "integration"):
        assert jobs[job]["if"] == expected, f"`{job}` if: {jobs[job].get('if')!r}"


def test_pr_size_gate_runs_in_pr_event_mode(workflow: dict) -> None:
    """Issue #857, PR-C: the size gate reads the PR event, not the runner state.

    The local mode depends on a `git diff` against `origin/main`, which is only
    present on a checked-out runner. Switching to `pr-event` lets the gate
    answer the same question from `$GITHUB_EVENT_PATH` alone.
    """
    steps = [s for s in workflow["jobs"]["pr-gate"]["steps"] if s.get("name") == "pr size"]
    assert len(steps) == 1, "pr-gate must have exactly one `pr size` step"
    step = steps[0]
    assert step["run"] == 'python scripts/check_pr_size.py pr-event "$GITHUB_EVENT_PATH"'
    env = step.get("env") or {}
    assert "BASE_REF" not in env
    assert "PR_BODY" not in env


def test_pr_size_gate_gets_the_token_for_the_files_breakdown(workflow: dict) -> None:
    """Issue #918: the size gate reads the per-file breakdown from the PR Files API.

    It needs `pull-requests: read` (the job-level grant) and the runner token, so the
    step gets the token explicitly; without it the gate counts the whole PR.
    """
    job = workflow["jobs"]["pr-gate"]
    assert job["permissions"]["pull-requests"] == "read"
    step = [s for s in job["steps"] if s.get("name") == "pr size"][0]
    assert step["env"]["GITHUB_TOKEN"] == "${{ github.token }}"


def test_issue_spec_is_wired_into_pr_gate(workflow: dict) -> None:
    """Issue #857, PR-C: the issue-spec gate reads the PR event and the linked issue.

    It needs `issues: read` (the job-level grant) and the runner token, so the
    step gets both explicitly rather than inheriting a broader default.
    """
    job = workflow["jobs"]["pr-gate"]
    assert job["permissions"] == {"contents": "read", "issues": "read", "pull-requests": "read"}
    steps = [step for step in job["steps"] if step.get("name") == "issue spec"]
    assert len(steps) == 1, "pr-gate must have exactly one `issue spec` step"
    step = steps[0]
    assert step["run"] == (
        'python scripts/check_issue_spec.py pr-event "$GITHUB_EVENT_PATH" '
        '--required-sections scripts/required-sections.json'
    )
    assert step["env"]["GITHUB_TOKEN"] == "${{ github.token }}"


def test_issue_spec_leaves_known_pending_gates(run_blocks: list[str]) -> None:
    """Issue #857, PR-C: once wired, check_issue_spec.py is no longer an exception."""
    assert "check_issue_spec.py" not in KNOWN_PENDING_GATES


def test_release_mutation_job_needs_preflight() -> None:
    """Issue #857, PR-C1d: the mutation job in release.yml runs after preflight.

    `preflight` is the gate that validates the tag identity, the ci verdict
    on the exact SHA, and the channel classification (rc vs stable). The
    mutation job depends on it because a failed preflight means the tag
    is not fit to ship, and running cosmic-ray on bytes that will never
    publish wastes the hosted runner's free-tier minutes.
    """
    release = yaml.safe_load((_find_workflow().parent / "release.yml").read_text(encoding="utf-8"))
    jobs = release["jobs"]
    assert "mutation" in jobs, "release.yml must declare a `mutation` job (#857, PR-C1d)"
    assert jobs["mutation"]["needs"] == "preflight", (
        f"`mutation` must declare `needs: preflight` (issue #857, PR-C1d). "
        f"Got: {jobs['mutation'].get('needs')!r}"
    )


def test_release_publish_needs_e2e_and_mutation() -> None:
    """Issue #857, PR-C1d: publish depends on both e2e and mutation.

    A green e2e is necessary (the published image must pass the Playwright
    smoke) but not sufficient — a mutation failure on the exact bytes
    being shipped is exactly the class of release-time failure the
    `mutation` job exists to catch. `publish` must wait for both, in
    either order, so a red mutation holds the tag.
    """
    release = yaml.safe_load((_find_workflow().parent / "release.yml").read_text(encoding="utf-8"))
    jobs = release["jobs"]
    assert "publish" in jobs, "release.yml must declare a `publish` job"
    needs = jobs["publish"]["needs"]
    # GitHub Actions serialises a list-typed `needs` as a list of strings;
    # the older scalar form is a bare string. Accept both.
    needed = {needs} if isinstance(needs, str) else set(needs)
    assert needed == {"e2e", "mutation"}, (
        f"`publish` must depend on `e2e` and `mutation` (issue #857, PR-C1d). "
        f"Got: {sorted(needed)!r}"
    )


def test_branch_name_is_no_longer_wired_into_ci(run_blocks: list[str]) -> None:
    """Issue #857, PR-B: the branch-name check leaves ci.yml; #859 deletes the script."""
    assert not [block for block in run_blocks if "check_branch_name.py" in block]


def test_release_signs_and_verifies_the_published_digest_with_oidc() -> None:
    """The release identity is the signed digest, not a mutable image tag."""
    release = yaml.safe_load((_find_workflow().parent / "release.yml").read_text(encoding="utf-8"))
    publish = release["jobs"]["publish"]
    verify = release["jobs"]["verify"]
    installer = "sigstore/cosign-installer@6f9f17788090df1f26f669e9d70d6ae9567deba6"

    assert publish["permissions"]["id-token"] == "write"
    assert "id-token" not in verify["permissions"]
    assert publish["outputs"]["digest"] == "${{ steps.push.outputs.digest }}"
    publish_installer = next(step for step in publish["steps"] if step.get("uses") == installer)
    verify_installer = next(step for step in verify["steps"] if step.get("uses") == installer)
    assert publish_installer["with"]["cosign-release"] == "v3.1.3"
    assert verify_installer["with"]["cosign-release"] == "v3.1.3"

    sign = next(step for step in publish["steps"] if step.get("name", "").startswith("Sign "))
    signature_check = next(
        step for step in verify["steps"] if step.get("name", "").startswith("Verify signature")
    )
    assert sign["env"]["DIGEST"] == "${{ steps.push.outputs.digest }}"
    assert sign["run"] == 'cosign sign --yes "$DIGEST"'
    assert signature_check["env"]["DIGEST"] == "${{ needs.publish.outputs.digest }}"
    assert 'cosign verify "$DIGEST"' in signature_check["run"]
    assert (
        '--certificate-identity "https://github.com/${GITHUB_REPOSITORY}/.github/workflows/'
        'release.yml@refs/tags/${GITHUB_REF_NAME}"'
    ) in signature_check["run"]
    assert (
        '--certificate-oidc-issuer "https://token.actions.githubusercontent.com"'
        in signature_check["run"]
    )


#: Herramientas que no provee el install del job `mutation`: las pone el runner
#: (`python`, `git`) o son construcciones del shell. Cualquier otro primer token
#: de un `run:` del job tiene que estar declarado en las extras `dev` del
#: proyecto y presente en su lock, porque es el install del job quien lo trae.
#: (#785: el paso `run mutants` moría con exit 127 —`cosmic-ray: command not
#: found`— porque el pin vivía en el asset de la skill y nunca se consolidó en
#: `app/pyproject.toml`.)
RUNNER_PROVIDED_BINARIES = frozenset(
    {
        "python",
        "python3",
        "git",
        "mkdir",
        "echo",
        "exit",
        "test",
        "set",
        "if",
        "then",
        "elif",
        "else",
        "fi",
        ":",
    }
)

_FIRST_TOKEN = re.compile(r"^\s*([A-Za-z0-9_./$@{}\-]+)")


def _invoked_binaries(run_blocks: list[str]) -> set[str]:
    """Primer token de cada línea de comando; sin asignaciones ni comentarios.

    Las continuaciones de línea (barra invertida final) se unen antes del primer token: una
    bandera de la línea siguiente no es un binario.
    """
    tokens: set[str] = set()
    for block in run_blocks:
        logical: list[str] = []
        for line in block.splitlines():
            if logical and logical[-1].endswith("\\"):
                logical[-1] = logical[-1][:-1].rstrip() + " " + line.strip()
            else:
                logical.append(line)
        for line in logical:
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            first = stripped.split()[0]
            if "=" in first:  # asignación (`blob_oid=$(...)`)
                continue
            match = _FIRST_TOKEN.match(stripped)
            if match is None:
                continue
            token = match.group(1)
            if token.startswith("${{") or token.endswith("/bin/python3"):
                token = "python"
            tokens.add(token)
    return tokens


#: Dominio declarado del job `mutation`: los dos directorios `domain/` del
#: monorepo. Medido el 2026-10-07: mutar `app/` entero genera 7905 mutantes y la
#: sesión no cabe en el techo de 6 h de un job hospedado; el dominio son 1129.
MUTATION_DOMAIN_TARGETS = (
    "app/src/modules/expedientes/domain",
    "app/src/modules/lanzadera/domain",
)

#: Áreas del árbol que quedan fuera del objetivo. Con la lista de arriba ya no
#: están en juego; se declaran para que ampliar `module-path` no las meta en
#: silencio (migraciones, el plugin de cobertura del arnés y el composition root).
MUTATION_EXCLUDED_AREAS = ("app/migrations/", "app/pytest_plugin/", "/di/")


def test_mutation_config_targets_only_the_domain() -> None:
    """#785: el objetivo del job `mutation` es el dominio, con timeout real.

    El config del PR apuntaba a `app` entero y a 30 s por mutante. La suite del
    `test-command` tarda 39,7-95,5 s en esta máquina, así que un superviviente se
    cortaba y cosmic-ray lo marcaba INCOMPETENT en vez de SURVIVED. Este test ata
    las tres piezas: el objetivo, las exclusiones y el timeout.
    """
    root = _find_makefile().parent
    config = tomllib.loads(
        (root / "docs" / "quality" / "cosmic-ray.toml").read_text(encoding="utf-8")
    )
    ray = config["cosmic-ray"]

    targets = ray.get("module-path")
    assert isinstance(targets, list), (
        "`module-path` debe ser la lista explícita de los directorios del dominio: "
        "una ruta ancha (p. ej. `app`) mete migraciones, el plugin de cobertura y "
        "el DI, y la sesión no cabe en un job (#785)"
    )
    assert tuple(targets) == MUTATION_DOMAIN_TARGETS, (
        f"el objetivo declarado es {MUTATION_DOMAIN_TARGETS}, no {tuple(targets)}"
    )
    for rel in targets:
        assert (root / rel).is_dir(), f"el objetivo declarado no existe: {rel}"

    excluded = ray.get("excluded-modules") or []
    for area in MUTATION_EXCLUDED_AREAS:
        assert any(area in pattern for pattern in excluded), (
            f"`excluded-modules` no declara el área {area!r}: ampliar `module-path` "
            "la metería en el objetivo sin que ningún gate lo note"
        )

    assert float(ray["timeout"]) >= 120.0, (
        f"timeout por mutante {ray['timeout']!r}: la suite del `test-command` tarda "
        "39,7-95,5 s, así que por debajo de ~120 s un superviviente se marca "
        "INCOMPETENT y el run sale degenerado (#785)"
    )


# ---------------------------------------------------------------------------
# Issue #857 (PR-A): trigger surface and mutation scope
# ---------------------------------------------------------------------------
#
# PR-A in the chain carries the trigger set and the mutation-scope change.
# PR-B carries the pr-gate rename, the needs wiring, the format --check
# migration, the sparse-checkout cleanup, and the check_required_jobs /
# required-jobs-policy / check_workflows renames. Each PR targets main;
# the two together close #857.


def test_ci_runs_on_push_to_main(workflow: dict) -> None:
    """Issue #857, PR-D: `push: branches: [main]` is the post-merge safety net.

    The gentle-ai model does not use a merge queue. Every PR runs `ci`
    exactly once, and the `required` aggregator re-proves the same
    commit on the push that follows the merge. The trigger is
    `branches: [main]` so a feature-branch push does not re-run the
    whole suite — only the merge to `main` does.
    """
    triggers = workflow.get(True) or workflow.get("on") or {}
    assert "push" in triggers, (
        "ci.yml must declare a `push:` trigger (issue #857, PR-D): the "
        "post-merge safety net on `main` is what re-proves the commit "
        "after the queue is gone."
    )
    push = triggers["push"]
    assert isinstance(push, dict), (
        f"ci.yml `push:` must declare `branches:` (issue #857, PR-D). Got: {push!r}"
    )
    assert list(push.get("branches", [])) == ["main"], (
        f"ci.yml `push:` must pin `branches: [main]` (issue #857, PR-D). "
        f"Got: {push.get('branches')!r}"
    )


def test_ci_runs_on_pull_request_push_schedule_dispatch(
    workflow: dict,
) -> None:
    """The four events the gentle-ai model surfaces (PR + push + nightly + manual)."""
    triggers = workflow.get(True) or workflow.get("on") or {}
    if isinstance(triggers, list):
        present = set(triggers)
    elif isinstance(triggers, str):
        present = {triggers}
    else:
        present = set(triggers.keys())

    for required_event in ("pull_request", "push", "schedule", "workflow_dispatch"):
        assert required_event in present, (
            f"ci.yml must declare `{required_event}` in its trigger set "
            f"(issue #857, PR-D). Present: {sorted(present)}."
        )


def test_ci_does_not_listen_to_merge_group(workflow: dict) -> None:
    """Issue #857, PR-D: the merge queue is gone with the gentle-ai model.

    `merge_group` was the queue's verdict event; without the queue there
    is no merge-group run to listen to. Leaving the trigger in place
    would re-introduce a class of `unsupported event` failures in
    `scripts/check_required_jobs.py` (#857 PR-D dropped it from
    `SKIPS_BY_EVENT`).
    """
    triggers = workflow.get(True) or workflow.get("on") or {}
    if isinstance(triggers, dict):
        assert "merge_group" not in triggers, (
            "ci.yml must not declare a `merge_group:` trigger: the queue is "
            "gone with the gentle-ai model (issue #857, PR-D)."
        )


# --- Issue #857, PR-B: pull_request type set -----------------------------
#
# `pull_request:` without an explicit `types:` uses the GitHub default
# (opened, synchronize, reopened). That set misses two events that DO
# invalidate an already-published pr-gate verdict:
#
#   * `edited`     — when the maintainer pastes `Closes #N` into the PR
#                    body after opening, the issue-spec gate does not
#                    re-run.
#   * `labeled` /
#     `unlabeled` — when the maintainer applies `size:exception` after
#                    the gate went red, the size verdict does not
#                    re-evaluate.
#
# Pinning the type set here is what guarantees a verdict refresh on
# those events. The list MUST stay exact (the order is the contract
# `check_workflows.py` proves).


PR_GATE_PULL_REQUEST_TYPES = (
    "opened",
    "synchronize",
    "reopened",
    "edited",
    "labeled",
    "unlabeled",
)


def test_pull_request_types_cover_edit_and_label_events(workflow: dict) -> None:
    """Issue #857, PR-B: pr-gate must re-run when body edits or labels change.

    The default ``pull_request`` trigger only fires on opened,
    synchronize, reopened. Editing the PR body or applying/removing
    ``size:exception`` after opening does NOT re-run the pr-gate, so a
    stale verdict sticks until the next synchronize. Pinning the type
    list here is what closes that drift.
    """
    # yaml.safe_load maps the `on` keyword to the Python boolean `True`
    # unless quoted; the existing tests guard both spellings already.
    triggers = workflow.get(True) or workflow.get("on") or {}
    pull_request = triggers.get("pull_request")
    assert isinstance(pull_request, dict), (
        "ci.yml must declare `pull_request:` with a body so `types:` "
        "can be pinned (issue #857, PR-B). A bare `pull_request:` falls "
        "back to the GitHub default type set and misses `edited`, "
        "`labeled` and `unlabeled`."
    )
    types = pull_request.get("types")
    assert types is not None, (
        "ci.yml must declare `pull_request.types:` — relying on the "
        "default leaves edit and label events out of the pr-gate "
        "(issue #857, PR-B)."
    )
    assert list(types) == list(PR_GATE_PULL_REQUEST_TYPES), (
        f"ci.yml `pull_request.types:` is {list(types)!r}, expected "
        f"{list(PR_GATE_PULL_REQUEST_TYPES)!r}. The set MUST include "
        "`edited`, `labeled` and `unlabeled` so the pr-gate re-runs "
        "when the body or labels change after opening (issue #857, PR-B)."
    )
