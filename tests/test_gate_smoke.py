# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — assets/tests/test_gate_smoke.py
"""Smoke tests: every gate is observed exiting 1 on a real violation (Execution Step 5).

A gate that has never been seen failing is a false guarantee. Each test here feeds a known
violation to a gate and asserts the failure, then feeds clean input and asserts the pass. When a
refactor silently breaks a detector, these go red before the detector's absence does damage.
"""

from __future__ import annotations

import importlib.util
import json
import re
import sqlite3
import subprocess
import sys
import tomllib
from pathlib import Path

import pytest

# Issue #882: the suite-populated guard (`scripts/check_suites_populated.py`)
# pins each mandatory suite against disappearing under a rename. Marking this
# file `e2e_smoke` is the way to satisfy the "≥1 test collects under
# `-m e2e_smoke`" side of the guard — without it, the guard fails
# immediately on a clean tree.
pytestmark = pytest.mark.e2e_smoke

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
FIXTURES = Path(__file__).resolve().parent / "fixtures"
#: DQH 2.2: check_complexity evaluates its BASELINE at an explicit date, never the wall clock.
#: check_mutation_sites, check_mutation and check_dry joined it when BP took the
#: canon's copies (BP #1316, #1327).
POLICY_DATE = "2026-10-21"
DATED_SCRIPTS = frozenset(
    {
        "check_complexity.py",
        "check_mutation_sites.py",
        "check_mutation.py",
        "check_dry.py",
    }
)


def _run(script: str, *args: str) -> subprocess.CompletedProcess[str]:
    if script in DATED_SCRIPTS:
        args = (*args, "--policy-date", POLICY_DATE)
    return subprocess.run(
        [sys.executable, str(SCRIPTS / script), *args],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )


def _load(script: str):
    name = script.removesuffix(".py")
    # The canon's gates import their siblings (`policy_time`, `quality_policy`) by module name
    # (DQH 2.2), so `scripts/` has to be importable before the file is executed.
    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))
    spec = importlib.util.spec_from_file_location(name, SCRIPTS / script)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    # Register before executing: dataclasses resolve their annotations through sys.modules, so
    # a module that is not registered cannot define a dataclass with a union annotation.
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


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
# check_dry
# ---------------------------------------------------------------------------------------------


def test_dry_gate_fails_on_a_renamed_copy_paste() -> None:
    result = _run("check_dry.py", "--root", str(FIXTURES / "dry_violation"))
    assert result.returncode == 1, result.stdout
    assert "alpha.py" in result.stdout and "beta.py" in result.stdout


def test_dry_gate_passes_on_distinct_code() -> None:
    result = _run("check_dry.py", "--root", str(FIXTURES / "dry_clean"))
    assert result.returncode == 0, result.stdout


def test_dry_baseline_key_survives_the_clone_spreading_to_another_file() -> None:
    """A ratchet whose key moves with the thing it ratchets is not a ratchet.

    The key used to be the joined file list, so copying the same block into one more file
    changed the key, the BASELINE entry stopped matching, and a grown clone read as a brand-new
    one. The key is the clone's own digest now: same duplication, same key, higher count.
    """
    module = _load("check_dry.py")
    two = module.collect_groups(FIXTURES / "dry_violation")
    three = module.collect_groups(FIXTURES / "dry_three_copies")
    assert two and three
    assert {group.key for group in two} == {group.key for group in three}
    assert len(three[0].occurrences) == 3, "the third copy must raise the count"


def test_dry_reports_a_long_clone_region_once() -> None:
    """A sliding window over one duplicated region produced one finding per offset.

    On a real codebase that turned 37 genuine clones into 333 findings — the exact noise level
    at which a team switches the gate off. The fixture's duplicated body is 6 statements, so a
    5-statement window slides over it twice; both offsets describe the same duplication and
    only one survives. A region long enough to hold two DISJOINT windows still reports two,
    which is correct: that really is two separate blocks of duplicated code.
    """
    module = _load("check_dry.py")
    groups = module.collect_groups(FIXTURES / "dry_long_region")
    assert len(groups) == 1, [group.key for group in groups]
    assert len(groups[0].occurrences) == 2


def test_dry_detection_is_stable_across_processes() -> None:
    """Determinism pin: the digest must not depend on PYTHONHASHSEED.

    An earlier draft keyed clone groups on the builtin ``hash()``, which is salted per process.
    The gate then reported different groups on identical code from one run to the next.
    """
    first = _run("check_dry.py", "--root", str(FIXTURES / "dry_violation"), "--json")
    second = _run("check_dry.py", "--root", str(FIXTURES / "dry_violation"), "--json")
    assert first.stdout == second.stdout


# ---------------------------------------------------------------------------------------------
# check_mutation_sites
# ---------------------------------------------------------------------------------------------


def test_mutation_sites_gate_fails_on_a_large_surface() -> None:
    result = _run("check_mutation_sites.py", "--root", str(FIXTURES / "mutation_sites_violation"))
    assert result.returncode == 1, result.stdout


def test_mutation_sites_catches_what_complexity_and_crap_cannot() -> None:
    """Surface is a property of the file, not of any one function.

    The fixture is all trivial functions and literals: complexity passes it, and CRAP would too.
    Only the site count says that changing anything in there is expensive.
    """
    root = FIXTURES / "mutation_sites_violation"
    assert _run("check_complexity.py", "--root", str(root)).returncode == 0
    assert _run("check_mutation_sites.py", "--root", str(root)).returncode == 1


def test_mutation_sites_gate_passes_on_a_small_surface() -> None:
    result = _run("check_mutation_sites.py", "--root", str(FIXTURES / "mutation_sites_clean"))
    assert result.returncode == 0, result.stdout


def test_mutation_sites_emits_a_baseline_with_target_and_date() -> None:
    """A ratchet you have to hand-write is a ratchet nobody adopts — and every emitted entry
    must already carry its exit plan, so an expiry-less baseline cannot be produced by accident.
    """
    result = _run(
        "check_mutation_sites.py",
        "--root",
        str(FIXTURES / "mutation_sites_violation"),
        "--emit-baseline",
    )
    assert result.returncode == 0
    assert "BaselineEntry(sites=" in result.stdout
    assert "target=" in result.stdout and "target_date=" in result.stdout


# ---------------------------------------------------------------------------------------------
# check_mutation
# ---------------------------------------------------------------------------------------------


#: DQH 2.2 (BP #1316): the session subjects must be inside `quality-policy.json`, so the
#: fixture uses a file the policy really measures instead of an invented path.
SUBJECT = "app/src/main.py"
#: A second real subject, for the two-session shard cases.
SECOND_SUBJECT = "app/src/modules/lanzadera/domain/user.py"


def _make_session(path: Path, rows: list[tuple[str, str | None]]) -> Path:
    """Build a minimal cosmic-ray session so these tests run without the mutation runner."""
    connection = sqlite3.connect(path)
    connection.executescript(
        """
        CREATE TABLE work_items (job_id TEXT PRIMARY KEY);
        CREATE TABLE mutation_specs (job_id TEXT, module_path TEXT,
                                     operator_name TEXT, occurrence INTEGER);
        CREATE TABLE work_results (job_id TEXT, test_outcome TEXT, worker_outcome TEXT);
        """
    )
    for index, (module, outcome) in enumerate(rows):
        job = f"job-{index}"
        connection.execute("INSERT INTO work_items VALUES (?)", (job,))
        connection.execute(
            "INSERT INTO mutation_specs VALUES (?, ?, ?, ?)", (job, module, "op", index)
        )
        if outcome is not None:
            connection.execute(
                "INSERT INTO work_results VALUES (?, ?, ?)", (job, outcome, "normal")
            )
    connection.commit()
    connection.close()
    return path


def test_mutation_gate_passes_a_healthy_run(tmp_path) -> None:
    session = _make_session(tmp_path / "s.sqlite", [(SUBJECT, "killed")] * 10)
    result = _run("check_mutation.py", str(session))
    assert result.returncode == 0, result.stdout + result.stderr


def test_mutation_gate_rejects_a_run_that_measured_nothing(tmp_path) -> None:
    """Hard Rule 18, pinned.

    Every mutant came back INCOMPETENT: the runner never executed. A plain survival-rate
    threshold reports this as a flawless run, which is the failure mode this gate exists for.
    """
    session = _make_session(tmp_path / "s.sqlite", [(SUBJECT, "incompetent")] * 10)
    result = _run("check_mutation.py", str(session))
    assert result.returncode == 1
    assert "degenerate" in result.stdout.lower()


def test_mutation_gate_rejects_an_incomplete_run(tmp_path) -> None:
    rows = [(SUBJECT, "killed")] * 5 + [(SUBJECT, None)] * 5
    session = _make_session(tmp_path / "s.sqlite", rows)
    result = _run("check_mutation.py", str(session))
    assert result.returncode == 1
    assert "incomplete" in result.stdout.lower()


def test_mutation_gate_fails_on_an_unrecorded_survivor(tmp_path) -> None:
    rows = [(SUBJECT, "killed")] * 9 + [(SUBJECT, "survived")]
    session = _make_session(tmp_path / "s.sqlite", rows)
    result = _run("check_mutation.py", str(session))
    assert result.returncode == 1
    assert "not in BASELINE" in result.stdout


def test_mutation_gate_fails_closed_on_a_missing_session(tmp_path) -> None:
    result = _run("check_mutation.py", str(tmp_path / "absent.sqlite"))
    assert result.returncode == 1


def test_mutation_gate_reports_score_and_incompetent_ratio(tmp_path) -> None:
    rows = [(SUBJECT, "killed")] * 8 + [(SUBJECT, "incompetent")]
    session = _make_session(tmp_path / "s.sqlite", rows)
    envelope = json.loads(_run("check_mutation.py", str(session), "--json").stdout)
    assert envelope["indicators"]["mutation_score_pct"] == 100.0
    assert envelope["indicators"]["incompetent_ratio_pct"] > 0


def test_acquisition_marker_cannot_become_permanent() -> None:
    """A module 'we have not measured yet' must stop being acceptable after the grace period."""
    module = _load("check_mutation.py")
    today = module.date(2026, 8, 8)
    module.BASELINE = {"app/fresh.py": module.AwaitingAcquisition(since="2026-08-05")}
    assert module.check_pending_overdue(today) == []
    module.BASELINE = {"app/stale.py": module.AwaitingAcquisition(since="2026-06-01")}
    assert module.check_pending_overdue(today), "an overdue acquisition marker must fail the gate"


# ---------------------------------------------------------------------------------------------
# quality_report — the indicator layer
# ---------------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "gate_script",
    [
        "check_complexity.py",
        "check_dry.py",
    ],
)
def test_every_gate_emits_a_well_formed_envelope(gate_script: str) -> None:
    root = FIXTURES / "complexity_clean"
    args = ["--root", str(root)]
    result = _run(gate_script, *args, "--json")
    envelope = json.loads(result.stdout)
    assert envelope["gate"]
    assert envelope["status"] in {"pass", "fail", "error"}
    assert isinstance(envelope["indicators"], dict)
    assert isinstance(envelope["findings"], list)


def test_quality_report_aggregates_indicators(tmp_path) -> None:
    out = tmp_path / "quality-report.json"
    result = _run(
        "quality_report.py", "--root", str(FIXTURES / "complexity_clean"), "--out", str(out)
    )
    assert result.returncode == 0, result.stdout
    report = json.loads(out.read_text(encoding="utf-8"))
    assert report["status"] == "pass"
    for indicator in (
        "complexity.max_complexity",
        "dry.duplicated_ratio_pct",
    ):
        assert indicator in report["indicators"], report["indicators"].keys()
    assert report["indicators"]["complexity.functions_over_ceiling"]["ceiling"] == 0


def test_quality_report_names_the_failing_gate(tmp_path) -> None:
    out = tmp_path / "quality-report.json"
    result = _run(
        "quality_report.py",
        "--root",
        str(FIXTURES / "dry_violation"),
        "--out",
        str(out),
    )
    assert result.returncode == 1
    report = json.loads(out.read_text(encoding="utf-8"))
    assert report["failed_gates"] == ["dry"]


def test_quality_report_is_byte_identical_for_the_same_commit(tmp_path) -> None:
    """Determinism pin: no wall-clock timestamp may leak into the report."""
    first, second = tmp_path / "a.json", tmp_path / "b.json"
    _run("quality_report.py", "--root", str(FIXTURES / "complexity_clean"), "--out", str(first))
    _run(
        "quality_report.py",
        "--root",
        str(FIXTURES / "complexity_clean"),
        "--out",
        str(second),
    )
    assert first.read_bytes() == second.read_bytes()


# --- Reparto del dominio entre los jobs de la matriz (#785) ---------------------------------
#
# El dominio (1129 mutantes) no cabe en un job. La matriz lo parte en tramos
# deterministas y el job final fusiona las sesiones.

MUTATION_DOMAIN_COUNTS = {
    "app/src/modules/expedientes/domain/anexo/retention.py": 116,
    "app/src/modules/expedientes/domain/anexo/size_limit.py": 99,
    "app/src/modules/expedientes/domain/cpv/code.py": 67,
    "app/src/modules/lanzadera/domain/legacy_role_map.py": 66,
    "app/src/modules/expedientes/domain/anualidad.py": 66,
    "app/src/modules/expedientes/domain/hash/versioning.py": 62,
}


def test_mutation_shard_assignment_is_complete_balanced_and_stable() -> None:
    """El reparto cubre el dominio una vez, reparte la carga y no depende del orden de entrada."""
    module = _load("mutation_shards.py")
    total = sum(MUTATION_DOMAIN_COUNTS.values())
    plan = module.assign(MUTATION_DOMAIN_COUNTS, 4)
    declared = [path for entry in plan for path in entry["modules"]]
    assert sorted(declared) == sorted(MUTATION_DOMAIN_COUNTS)
    assert len(declared) == len(set(declared))
    assert [entry["index"] for entry in plan] == [0, 1, 2, 3]
    loads = [entry["mutants"] for entry in plan]
    assert max(loads) <= total / len(plan) * 1.2, loads
    reversed_input = dict(reversed(list(MUTATION_DOMAIN_COUNTS.items())))
    assert module.assign(reversed_input, 4) == plan, (
        "el reparto no puede depender del orden de entrada"
    )


def test_mutation_shard_assignment_fails_closed_with_too_few_modules() -> None:
    """Un módulo no se parte entre dos jobs: con menos módulos que tramos, error."""
    module = _load("mutation_shards.py")
    with pytest.raises(module.ShardError):
        module.assign({SUBJECT: 3}, 4)
    with pytest.raises(module.ShardError):
        module.assign(MUTATION_DOMAIN_COUNTS, 0)


def test_mutation_shard_config_keeps_the_shared_config() -> None:
    """El config del tramo no puede divergir del compartido en timeout ni en test-command."""
    module = _load("mutation_shards.py")
    shared = tomllib.loads(
        (ROOT / "docs" / "quality" / "cosmic-ray.toml").read_text(encoding="utf-8")
    )["cosmic-ray"]
    rendered = module.render(shared, ["app/src/modules/lanzadera/domain/presence.py"])
    parsed = tomllib.loads(rendered)["cosmic-ray"]
    assert parsed["module-path"] == ["app/src/modules/lanzadera/domain/presence.py"]
    assert parsed["timeout"] == shared["timeout"]
    assert parsed["test-command"] == shared["test-command"]
    assert parsed["excluded-modules"] == shared["excluded-modules"]


def test_mutation_gate_accepts_one_session_per_shard(tmp_path) -> None:
    first = _make_session(tmp_path / "s0.sqlite", [(SUBJECT, "killed")] * 5)
    second = _make_session(tmp_path / "s1.sqlite", [(SECOND_SUBJECT, "killed")] * 5)
    result = _run("check_mutation.py", str(first), str(second), "--shards", "2")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "run healthy" in result.stdout


def test_mutation_gate_fails_closed_when_a_shard_is_missing(tmp_path) -> None:
    first = _make_session(tmp_path / "s0.sqlite", [(SUBJECT, "killed")] * 5)
    second = _make_session(tmp_path / "s1.sqlite", [(SECOND_SUBJECT, "killed")] * 5)
    result = _run("check_mutation.py", str(first), str(second), "--shards", "3")
    assert result.returncode == 1
    assert "se esperaban 3" in result.stderr


def test_mutation_gate_fails_closed_when_a_shard_is_degenerate(tmp_path) -> None:
    healthy = _make_session(tmp_path / "s0.sqlite", [(SUBJECT, "killed")] * 5)
    degenerate = _make_session(tmp_path / "s1.sqlite", [("app/src/main.py", "incompetent")] * 5)
    result = _run("check_mutation.py", str(healthy), str(degenerate), "--shards", "2")
    assert result.returncode == 1
    assert "degenerate" in result.stdout.lower()
    assert "s1.sqlite" in result.stdout


def test_mutation_test_command_caps_memory_per_run() -> None:
    """El `test-command` limita la memoria de cada ejecución (#785).

    Un mutante desbocado mató cuatro veces el runner hospedado (SIGTERM) antes del
    timeout de 120 s: el que muere es el runner, no el test. El tope hace que ese
    mutante falle el test y cuente como «killed». Medido: la suite normal pica en
    177 MB de RSS y pasa con 1 GiB; el config declara 2 GiB.
    """
    shared = tomllib.loads(
        (ROOT / "docs" / "quality" / "cosmic-ray.toml").read_text(encoding="utf-8")
    )["cosmic-ray"]
    command = shared["test-command"]
    match = re.search(r"ulimit -v (\d+)", command)
    assert match, f"el test-command no limita el espacio de direcciones: {command}"
    assert int(match.group(1)) >= 1048576, (
        f"tope {match.group(1)} KB: por debajo de 1 GiB no es holgado para una suite "
        "que pica en 177 MB (el valor declarado son 2 GiB)"
    )
    assert "--no-cov" in command
