# HARNESS-PROVENANCE: deterministic-quality-harness v1.5 — assets/tests/test_ci_workflow.py
"""Wiring pin for the CI workflow (Hard Rule 4).

A gate that exists only as a script is a gate nobody runs. These tests parse the workflow and
fail when a step is deleted, when a gate is neutered, or when a pin goes floating. They are the
reason a silent revert cannot pass review.

Issue #1375: the pins of the SHAPE of `ci.yml` — its event list, the exact `if:` of the heavy
jobs, their `needs:`, the `uses:` of the gate job — live in the canon's structural gates
(`check_workflow_policy.py`, `check_parity.py`) and in the required-jobs policy, and they
churned 37 times next to `ci.yml` in three weeks. What stays in this file pins what the gates
DO: the commands they run, how they are hardened, and that each gate exists and runs.
"""

from __future__ import annotations

import importlib.util
import json
import re
import sys
import tomllib
from pathlib import Path

import pytest
import yaml

#: Every gate that must be wired, identified by the command it runs.
#: La cadena de dependencias (HR-76, DysTelefonica/team-skills#1181): una sola
#: linea, la misma en el job pr-gate y en scripts/local-preflight.sh.
DEPS_CHAIN_COMMAND = (
    "git diff --name-only origin/${GITHUB_BASE_REF:-main}...HEAD"
    ' > "${RUNNER_TEMP:-/tmp}/changed-files.txt"'
    " && python3 .agents/skills/ci-pattern/assets/deps-chain/check_deps_chain.py"
    ' --repo . --changed-files "${RUNNER_TEMP:-/tmp}/changed-files.txt"'
)

REQUIRED_COMMANDS = (
    "ruff format --check --config app/pyproject.toml",
    "ruff check --config app/pyproject.toml .",
    # El preflight de tooling del pr-gate: en el workflow el comando ocupa tres
    # lineas (continuacion con `\`), asi que la firma que se puede comparar es
    # la del binario que el gate ejecuta.
    "sh .agents/skills/ci-pattern/assets/tooling-preflight/tooling-preflight.sh",
    "mypy --explicit-package-bases app/",
    "python scripts/check_workflows.py",
    'pytest -c app/pyproject.toml --rootdir=app -m "not integration"',
    "python scripts/check_workflows.py",
    "scripts/quality_report.py",
    DEPS_CHAIN_COMMAND,
    ".agents/skills/ci-pattern/assets/pr-size/check_pr_size.py",
    ".agents/skills/ci-pattern/assets/issue-spec/check_issue_spec.py",
    "python scripts/check_walkthrough_schema.py --strict",
    ".agents/skills/ci-pattern/assets/required-jobs/check_required_jobs.py",
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
    ".agents/skills/ci-pattern/assets/required-jobs/check_required_jobs.py",
)

#: The code gates, in the order they must run. Deduplication moves code, which changes
#: complexity — so the sequence is part of the contract, not a preference.
REQUIRED_GATE_ORDER = (
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
def pr_gate_workflow() -> dict:
    """The lightweight workflow that owns the `pr-gate` job (issue #1368).

    Since #1368 the `pr-gate` (body, labels, walkthrough) lives in its own
    workflow: the body/label events must not drag the heavy suite along. The
    heavy workflow calls this file, so `required` keeps aggregating the verdict.
    """
    path = _find_workflow().parent / "pr-gate.yml"
    assert path.is_file(), (
        "the pr-gate job lives in .github/workflows/pr-gate.yml since #1368: the "
        "edited/labeled/unlabeled events must not run the heavy suite"
    )
    return yaml.safe_load(path.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def steps(workflow: dict, pr_gate_workflow: dict) -> list[dict]:
    collected: list[dict] = []
    for doc in (workflow, pr_gate_workflow):
        for job in doc["jobs"].values():
            collected.extend(job.get("steps", []))
    return collected


@pytest.fixture(scope="module")
def run_blocks(steps: list[dict]) -> list[str]:
    return [step["run"] for step in steps if "run" in step]


@pytest.mark.parametrize("command", REQUIRED_COMMANDS)
def test_gate_is_wired(run_blocks: list[str], command: str) -> None:
    assert any(command in block for block in run_blocks), f"no CI step runs '{command}'"


def test_gate_order_is_pinned_in_code(monkeypatch) -> None:
    """Hard Rule 13: the order lives in quality_report.GATES, and this is what pins it.

    Keeping the four code gates inside one script rather than four YAML steps is deliberate: a
    reviewer can reorder two YAML steps without noticing, and the resulting verdict would depend
    on the order in which someone happened to fix things.
    """
    scripts = Path(__file__).resolve().parent.parent / "scripts" / "quality_report.py"
    # DQH 2.2: the gates import their siblings by module name, so `scripts/` has to be
    # importable for this file to run on its own. `syspath_prepend` and `setitem` both
    # restore what they touched, so this test stops depending on some other test having
    # polluted the interpreter first (BP#1353).
    monkeypatch.syspath_prepend(str(scripts.parent))
    spec = importlib.util.spec_from_file_location("quality_report", scripts)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    monkeypatch.setitem(sys.modules, "quality_report", module)
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


def test_pr_size_gate_runs_in_pr_event_mode(pr_gate_workflow: dict) -> None:
    """Issue #857, PR-C: the size gate reads the PR event, not the runner state.

    The local mode depends on a `git diff` against `origin/main`, which is only
    present on a checked-out runner. Switching to `pr-event` lets the gate
    answer the same question from `$GITHUB_EVENT_PATH` alone.
    """
    steps = [s for s in pr_gate_workflow["jobs"]["pr-gate"]["steps"] if s.get("name") == "pr size"]
    assert len(steps) == 1, "pr-gate must have exactly one `pr size` step"
    step = steps[0]
    assert step["run"] == (
        'python .agents/skills/ci-pattern/assets/pr-size/check_pr_size.py --event "$GITHUB_EVENT_PATH"'
    )
    env = step.get("env") or {}
    assert "BASE_REF" not in env
    assert "PR_BODY" not in env


def test_pr_size_gate_gets_the_token_for_the_files_breakdown(pr_gate_workflow: dict) -> None:
    """Issue #918: the size gate reads the per-file breakdown from the PR Files API.

    It needs `pull-requests: read` (the job-level grant) and the runner token, so the
    step gets the token explicitly; without it the gate counts the whole PR.
    """
    job = pr_gate_workflow["jobs"]["pr-gate"]
    assert job["permissions"]["pull-requests"] == "read"
    step = [s for s in job["steps"] if s.get("name") == "pr size"][0]
    assert step["env"]["GITHUB_TOKEN"] == "${{ github.token }}"


def test_workflow_policy_gate_runs_the_canon_over_every_workflow(workflow: dict) -> None:
    """BP#1040 (B3b): the canon workflow-policy gate covers every workflow file, not only ci.yml.

    The local `workflow gate` keeps only what the canon does not check; the rest (duplicate keys,
    timeouts, host ports, SHA pins, docker preflight, concurrency groups) is the canon's job, so
    the canon has to see every file the local gate used to scan.
    """
    steps = [s for s in workflow["jobs"]["quality"]["steps"] if s.get("name") == "workflow policy"]
    assert len(steps) == 1, "quality must have exactly one `workflow policy` step"
    run = steps[0]["run"]
    assert 'for f in .github/workflows/*.yml; do args+=(--workflow "$f"); done' in run
    assert (
        "python .agents/skills/ci-pattern/assets/workflow-policy/check_workflow_policy.py "
        '"${args[@]}" --base-branch main --required-jobs .github/required-jobs-policy.json'
    ) in run


def test_pip_audit_audits_the_hash_locked_set_strictly() -> None:
    """The dependency scan audits `app/requirements-dev.lock`, the set that is installed and shipped.

    `pip install "app[dev]"` is a registry requirement named `app` to pip (the path form needs a
    `./` prefix), so it audited a homonym on PyPI and never the project. The lock is audited
    directly and strictly, with hashes required and pip disabled, and an empty lock cannot pass.
    """
    workflows = _find_workflow().parent
    audit = yaml.safe_load((workflows / "pip-audit.yml").read_text(encoding="utf-8"))
    steps = [
        s for s in audit["jobs"]["dependency-scan"]["steps"] if s.get("name") == "dependency scan"
    ]
    assert len(steps) == 1, "pip-audit must have exactly one `dependency scan` step"
    run = steps[0]["run"]
    assert "app[dev]" not in run
    assert (
        ".venv-audit/bin/pip-audit --strict --require-hashes --disable-pip -r app/requirements-dev.lock"
        in run
    )
    assert "pip-audit==2.9.0" in run
    assert "-lt 15" in run, "an empty or truncated lock must fail before the verdict is trusted"

    # Issue #1377: the job is a first-level job of `ci.yml` (its own reusable
    # workflow), so the `required` aggregator can read its result and declare
    # its skip by design.
    workflow = yaml.safe_load(_find_workflow().read_text(encoding="utf-8"))
    job = workflow["jobs"]["pip-audit"]
    assert job["uses"] == "./.github/workflows/pip-audit.yml", (
        "the dependency scan must be a first-level job of ci.yml: `needs:` cannot cross "
        "workflow files, so an inner job of security.yml could not declare its own skip"
    )
    assert job["needs"] == "pr-gate"
    assert (
        "pip-audit"
        not in yaml.safe_load((workflows / "security.yml").read_text(encoding="utf-8"))["jobs"]
    ), "the dependency scan moved out of security.yml (issue #1377)"


def test_issue_spec_is_wired_into_pr_gate(pr_gate_workflow: dict) -> None:
    """Issue #857, PR-C: the issue-spec gate reads the PR event and the linked issue.

    It needs `issues: read` (the job-level grant) and the runner token, so the
    step gets both explicitly rather than inheriting a broader default. Issue #1330:
    `--host-contract` made the gate check who applied `status:approved`; that rule
    was retired (team-skills#1103), so the flag went with it and the gate no longer
    takes it (team-skills#1121 keeps accepting it as an obsolete no-op).
    """
    job = pr_gate_workflow["jobs"]["pr-gate"]
    assert job["permissions"] == {"contents": "read", "issues": "read", "pull-requests": "read"}
    steps = [step for step in job["steps"] if step.get("name") == "issue spec"]
    assert len(steps) == 1, "pr-gate must have exactly one `issue spec` step"
    step = steps[0]
    assert step["run"] == (
        'python .agents/skills/ci-pattern/assets/issue-spec/check_issue_spec.py pr-event "$GITHUB_EVENT_PATH"'
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


# --- Issue #857, PR-B: pull_request type set of the pr-gate ---------------
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
#
# Issue #1375: only the list of `pr-gate.yml` is pinned here. The event
# set of `ci.yml` is a cost decision, not a verdict: the canon's
# required-jobs gate owns the event surface of the workflow, and the
# shape of `ci.yml` is no longer pinned by this file (#1368 moved the
# body and label events to this workflow).


PR_GATE_PULL_REQUEST_TYPES = (
    "opened",
    "synchronize",
    "reopened",
    "edited",
    "labeled",
    "unlabeled",
)


def test_pr_gate_listens_to_edit_and_label_events(pr_gate_workflow: dict) -> None:
    """Issue #1368: the body/label events moved to the lightweight workflow.

    Issue #857, PR-B pinned `edited`, `labeled` and `unlabeled` on the
    workflow that also runs the heavy suite, so a body edit re-ran
    `quality`, `integration`, `security`, `codeql` and `e2e-smoke` over a SHA
    that was already green. Since #1368 the heavy workflow listens to the
    three events that can change the code, and the `pr-gate` — the only job
    those events need — lives in its own workflow, which still listens to all
    six. This list is pinned so a silent drift cannot pass review.
    """
    # yaml.safe_load maps the `on` keyword to the Python boolean `True`
    # unless quoted; the existing tests guard both spellings already.
    light_types = (pr_gate_workflow.get(True) or pr_gate_workflow.get("on") or {}).get(
        "pull_request"
    )
    assert isinstance(light_types, dict), (
        "pr-gate.yml must declare `pull_request:` with a body so `types:` can be pinned"
    )
    assert list(light_types.get("types") or []) == list(PR_GATE_PULL_REQUEST_TYPES), (
        f"pr-gate.yml `pull_request.types:` is {light_types.get('types')!r}, expected "
        f"{list(PR_GATE_PULL_REQUEST_TYPES)!r}. The set MUST include `edited`, "
        "`labeled` and `unlabeled` so the pr-gate re-runs when the body or labels "
        "change after opening (issue #857, PR-B; moved by #1368)."
    )


#: The jobs the docs-only lane skips: the heavy suite of the PR rollup.
HEAVY_JOBS = (
    "quality",
    "integration",
    "security",
    "codeql",
    "e2e-smoke",
)


def test_the_pr_gate_queue_separates_the_two_instances(pr_gate_workflow: dict) -> None:
    """BP#1372: the same gate runs in two instances — its own workflow and the call from `ci.yml`.

    A concurrency group holds one pending request: GitHub cancels the pending one when
    another arrives, so the instance called from `ci.yml` was cancelled at 0 s and the
    `required` aggregator read the cancellation as a failure (run `ci` 40084048493,
    attempt 1: "Canceling since a higher priority waiting request for gate-pr-gate-1371
    exists"). The group must name the instance: inside a reusable workflow
    `github.workflow` is the caller's name, so the two instances land in different queues
    and neither cancels the other.
    """
    group = pr_gate_workflow["jobs"]["pr-gate"]["concurrency"]["group"]
    assert "github.workflow" in group, (
        f"the pr-gate concurrency group must separate the instance, got {group!r}"
    )


def test_the_heavy_jobs_skip_the_docs_only_lane(workflow: dict) -> None:
    """Issue #1368: a diff entirely under `docs/` skips the heavy suite.

    The marker is the `docs_changed` output of the pr-gate. Every heavy job
    must read it, or the lane would skip some jobs and not others, and the
    `required` aggregator would see a half-skipped run.
    """
    for job in HEAVY_JOBS:
        condition = workflow["jobs"][job]["if"]
        assert "needs.pr-gate.outputs.docs_changed != 'true'" in condition, (
            f"`{job}` must skip the docs-only lane: {condition!r}"
        )
    # `docs-links` is the gate that catches broken links in the PRD; it stays
    # in the run when the diff is documentation.
    assert "docs_changed" not in workflow["jobs"]["docs-links"]["if"]


def test_the_policy_declares_the_docs_only_lane(workflow: dict) -> None:
    """The lane is declared in the policy, and it covers exactly the heavy jobs.

    A skip the policy does not declare leaves `required` in red, and a skip
    the policy declares for a job that never skips is a promise the gate
    cannot keep.
    """
    policy = json.loads(
        (_find_workflow().parent.parent / "required-jobs-policy.json").read_text(encoding="utf-8")
    )
    lane = policy["docs_only_lane"]
    assert lane["event"] == "pull_request"
    assert lane["gate_job"] == "pr-gate"
    assert lane["output"] == "docs_changed"
    assert sorted(lane["skips"]) == sorted(HEAVY_JOBS)
    for job in lane["skips"]:
        assert job in policy["required_jobs"], f"{job} is skipped but not a required job"
    assert "relaxation" in policy, "a new lane is a relaxation against the base policy"


def test_the_docs_only_classifier_runs_in_the_pr_gate(pr_gate_workflow: dict) -> None:
    """A gate that exists only as a script is a gate nobody runs (Hard Rule 4)."""
    runs = "\n".join(
        step.get("run", "") for step in pr_gate_workflow["jobs"]["pr-gate"].get("steps", [])
    )
    assert "scripts/check_docs_only_lane.py" in runs
    assert "docs_changed" in runs, "the classifier's verdict must reach the job output"


def test_docs_codeguide_job_covers_every_root_of_the_adoption_contract() -> None:
    """HR-63 (team-skills#805): the CI gate checks exactly the roots the contract declares."""
    import json

    root = Path(__file__).resolve().parents[1]
    doc = yaml.safe_load((root / ".github" / "workflows" / "ci.yml").read_text(encoding="utf-8"))
    run = next(
        step["run"]
        for step in doc["jobs"]["docs-codeguide"]["steps"]
        if "check_docs_codeguide.py" in step.get("run", "")
    )
    pairs = re.findall(r'--code-root (\S+)[\s\\]+--code-glob "([^"]+)"', run)
    contract = json.loads(
        (root / ".github" / "ci-pattern-adoption.json").read_text(encoding="utf-8")
    )
    declared = [(e["root"], e["glob"]) for e in contract["docs_map"]["code_roots"]]
    assert pairs == declared


def test_e2e_smoke_job_runs_on_pull_request_and_is_required(workflow: dict) -> None:
    """Issue #876: the app boots and answers on every PR, not only at release time.

    The job builds the PR image, waits up to 60 s for `/health` and runs the
    Playwright smoke specs. It runs only on `pull_request`; on the other events
    it is skipped and the policy accepts that skip. `release.yml` keeps its own
    full E2E, which this job does not replace.
    """
    jobs = workflow["jobs"]
    assert "e2e-smoke" in jobs, "ci.yml must declare the `e2e-smoke` job (#876)"
    job = jobs["e2e-smoke"]

    assert job["needs"] == "pr-gate"
    assert "github.event_name == 'pull_request'" in job["if"]
    assert "always()" not in job["if"]
    assert job["runs-on"] == "ubuntu-24.04", "hosted runner: fork PRs are untrusted input"
    assert job["timeout-minutes"] == 5
    assert "e2e-smoke" in jobs["required"]["needs"]

    steps = "\n".join(str(step.get("run", "")) for step in job["steps"])
    assert "docker build" in steps
    assert "seq 1 60" in steps and "/health" in steps, "/health must be awaited for 60 s"
    assert "playwright test smoke.spec.ts auth.spec.ts" in steps
    assert "|| true" not in steps and "continue-on-error" not in str(job)

    release = yaml.safe_load((_find_workflow().parent / "release.yml").read_text(encoding="utf-8"))
    assert "e2e" in release["jobs"], "the release E2E job stays as it is"


def test_nightly_pip_audit_failure_opens_an_issue(workflow: dict) -> None:
    # HR-39 (#1414): a red scheduled pip-audit must leave an issue with the run link,
    # and only a scheduled failure may open one.
    job = workflow["jobs"]["pip-audit-alert"]
    assert job["needs"] == "pip-audit"
    assert job["if"] == "failure() && github.event_name == 'schedule'"
    assert "gh issue create" in "\n".join(step.get("run", "") for step in job["steps"])


def test_only_the_pip_audit_alert_job_may_write_issues(workflow: dict) -> None:
    # HR-39 (#1414): `issues: write` is granted to the alert job alone.
    writers = sorted(
        name
        for name, job in workflow["jobs"].items()
        if (job.get("permissions") or {}).get("issues") == "write"
    )
    assert writers == ["pip-audit-alert"]
