"""BP's mutation gate reads its subjects and ceilings from ``quality-policy.json``.

The gate is the canon's asset (``deterministic-quality-harness`` v2.2) with two declared local
deviations: ``--shards`` over a session per matrix slice, and a BASELINE aggregated by module
instead of by mutant (re-acquiring it per mutant needs the ARM matrix run — BP #1316). These
tests pin the canon wiring that was adopted, the ratchet that was kept, and the shard contract
that ``release.yml`` depends on.
"""

from __future__ import annotations

import importlib.util
import json
import sqlite3
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
POLICY_DATE = "2026-10-21"

_SCHEMA = """
CREATE TABLE mutation_specs (job_id INTEGER PRIMARY KEY, module_path TEXT,
                             operator_name TEXT, occurrence INTEGER);
CREATE TABLE work_items (job_id INTEGER PRIMARY KEY);
CREATE TABLE work_results (job_id INTEGER PRIMARY KEY, test_outcome TEXT, worker_outcome TEXT);
"""


def _session(path: Path, rows: list[tuple[str, str, str, str]]) -> Path:
    """Write a cosmic-ray-shaped session: (module, operator, occurrence, outcome)."""
    connection = sqlite3.connect(path)
    try:
        connection.executescript(_SCHEMA)
        for index, (module, operator, occurrence, outcome) in enumerate(rows, start=1):
            connection.execute(
                "INSERT INTO mutation_specs VALUES (?, ?, ?, ?)",
                (index, module, operator, occurrence),
            )
            connection.execute("INSERT INTO work_items VALUES (?)", (index,))
            connection.execute(
                "INSERT INTO work_results VALUES (?, ?, ?)", (index, outcome, "completed")
            )
        connection.commit()
    finally:
        connection.close()
    return path


def _subject(root: Path) -> str:
    """A file that the quality policy really measures, so the subject is inside policy."""
    package = root / "app" / "pkg"
    package.mkdir(parents=True)
    (package / "surface.py").write_text(
        "def f(x: int) -> int:\n    return x + 1\n", encoding="utf-8"
    )
    return "app/pkg/surface.py"


def _load():
    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))
    spec = importlib.util.spec_from_file_location("check_mutation", SCRIPTS / "check_mutation.py")
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    # Registered before execution: the dataclasses resolve their annotations through sys.modules.
    sys.modules["check_mutation"] = module
    spec.loader.exec_module(module)
    return module


def _run(root: Path, session: Path, *extra: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            sys.executable,
            str(SCRIPTS / "check_mutation.py"),
            str(session),
            "--root",
            str(root),
            *extra,
        ],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )


def test_an_explicit_policy_date_is_recorded_in_the_envelope(tmp_path: Path) -> None:
    subject = _subject(tmp_path)
    session = _session(tmp_path / "s.sqlite", [(subject, "op", 1, "killed")])
    result = _run(tmp_path, session, "--policy-date", POLICY_DATE, "--json")
    assert json.loads(result.stdout)["policy_date"] == POLICY_DATE


def test_without_a_policy_date_the_gate_refuses_to_run(tmp_path: Path) -> None:
    """The BASELINE is judged at an explicit date: without one the gate measures nothing (BP #1324).

    Falling back to the wall clock would make the same session yield a different verdict tomorrow,
    which is the determinism the canon's `policy_time` exists to prevent.
    """
    subject = _subject(tmp_path)
    session = _session(tmp_path / "s.sqlite", [(subject, "op", 1, "killed")])
    result = _run(tmp_path, session)
    assert result.returncode == 2, result.stdout + result.stderr
    assert "--policy-date" in result.stderr


def test_the_baseline_is_judged_at_the_date_it_is_given(tmp_path: Path) -> None:
    """The explicit date drives the ratchet, never the wall clock (Hard Rule 12 exit plan)."""
    subject = _subject(tmp_path)
    session = _session(
        tmp_path / "s.sqlite",
        [(subject, "op", 1, "killed"), (subject, "op", 2, "survived")],
    )
    module = _load()
    module.BASELINE = {
        subject: module.BaselineEntry(survivors=1, target=0, target_date="2026-10-20")
    }
    argv = [str(session), "--root", str(tmp_path)]
    assert module.main([*argv, "--policy-date", "2026-10-20"]) == 0
    assert module.main([*argv, "--policy-date", POLICY_DATE]) == 1


def test_envelope_is_bound_to_the_policy_and_its_date(tmp_path: Path) -> None:
    subject = _subject(tmp_path)
    session = _session(
        tmp_path / "s.sqlite",
        [(subject, "op", 1, "killed"), (subject, "op", 2, "survived")],
    )
    result = _run(tmp_path, session, "--policy-date", POLICY_DATE, "--json")
    envelope = json.loads(result.stdout)
    assert envelope["policy_date"] == POLICY_DATE
    assert envelope["ceilings"] == {"incompetent_ratio_pct": 20.0}
    assert envelope["subjects"]["checked"] == [subject]
    assert envelope["policy_identity"].startswith("sha256:")
    assert envelope["indicators"]["survivors_total"] == 1
    assert envelope["indicators"]["mutants_measured"] == 2


def test_a_subject_outside_the_quality_policy_is_rejected(tmp_path: Path) -> None:
    _subject(tmp_path)
    session = _session(
        tmp_path / "s.sqlite",
        [("tools/thing.py", "op", 1, "killed"), ("tools/thing.py", "op", 2, "survived")],
    )
    result = _run(tmp_path, session, "--policy-date", POLICY_DATE)
    assert result.returncode == 1
    assert "subjects outside quality policy" in result.stdout
    assert "tools/thing.py" in result.stdout


def test_the_aggregate_baseline_still_judges_survivors(tmp_path: Path) -> None:
    """The kept deviation: survivors are counted per module, and an unknown module fails."""
    subject = _subject(tmp_path)
    session = _session(
        tmp_path / "s.sqlite",
        [(subject, "op", 1, "killed"), (subject, "op", 2, "survived")],
    )
    result = _run(tmp_path, session, "--policy-date", POLICY_DATE)
    assert result.returncode == 1
    assert f"{subject}: 1 surviving mutant(s), not in BASELINE" in result.stdout


def test_the_shard_count_is_still_enforced(tmp_path: Path) -> None:
    """``release.yml`` passes one session per slice; a missing slice must fail the gate."""
    subject = _subject(tmp_path)
    session = _session(tmp_path / "s1.sqlite", [(subject, "op", 1, "killed")])
    result = _run(tmp_path, session, "--shards", "2", "--policy-date", POLICY_DATE)
    assert result.returncode == 1
    assert "se esperaban 2 sesión(es) y llegaron 1" in result.stderr
