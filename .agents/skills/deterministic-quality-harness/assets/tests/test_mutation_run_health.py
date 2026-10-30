# HARNESS-PROVENANCE: deterministic-quality-harness v2.2 — assets/tests/test_mutation_run_health.py
"""A mutation run that killed nothing did not measure the code (#809 T2b).

``check_run_health`` already refuses a session with no jobs at all. What it did not refuse
was the session that produced rows and killed none of them: every survivor in that session
came from a suite that never ran against the mutated code, so the score it yields is noise.
"""

from __future__ import annotations

import contextlib
import importlib.util
import io
import sqlite3
import sys
from datetime import date
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
POLICY_DATE = "2026-08-12"


def _load():
    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))
    spec = importlib.util.spec_from_file_location("check_mutation", SCRIPTS / "check_mutation.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_mutation"] = module
    spec.loader.exec_module(module)
    return module


def _empty_session(tmp_path: Path) -> Path:
    path = tmp_path / "session.db"
    connection = sqlite3.connect(path)
    connection.executescript(
        """
        CREATE TABLE work_items (job_id TEXT PRIMARY KEY);
        CREATE TABLE mutation_specs (
            job_id TEXT, module_path TEXT, operator_name TEXT, occurrence INTEGER
        );
        CREATE TABLE work_results (job_id TEXT, test_outcome TEXT, worker_outcome TEXT);
        """
    )
    connection.commit()
    connection.close()
    return path


def _rows(*outcomes: str) -> list[dict[str, str]]:
    return [{"module_path": "app/x.py", "test_outcome": outcome} for outcome in outcomes]


def test_a_run_that_killed_nothing_is_rejected() -> None:
    module = _load()

    violations, _ = module.check_run_health(_rows(*([module._SURVIVED] * 12)))

    assert any("were killed" in violation for violation in violations), violations


def test_the_rejection_names_the_measured_total() -> None:
    """The message must carry N, so a reader can tell 3 mutants from 300."""
    module = _load()

    violations, _ = module.check_run_health(_rows(*([module._SURVIVED] * 3)))

    assert any("0/3" in violation for violation in violations), violations


def test_a_run_that_killed_something_is_accepted() -> None:
    module = _load()

    violations, _ = module.check_run_health(_rows(module._KILLED, *([module._SURVIVED] * 11)))

    assert violations == [], violations


def test_an_empty_session_keeps_its_existing_verdict() -> None:
    """Zero mutants is not the rule this change adds: the pre-existing refusal stands."""
    module = _load()

    violations, metrics = module.check_run_health([])

    assert violations == ["session contains no mutation jobs; refusing to report success"]
    assert metrics == {}


def _run_main(module, session: Path, root: Path) -> tuple[int, str]:
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):
        code = module.main([str(session), "--root", str(root), "--policy-date", POLICY_DATE])
    return code, buffer.getvalue()


def test_a_session_with_no_jobs_is_refused_and_says_why(tmp_path: Path) -> None:
    """D6: zero mutants is a NOTE beside the existing refusal, not a "0/0" failure.

    The rule this change adds covers a session that measured mutants and killed none. A
    session with no jobs was already refused, and the note states that the refusal is about
    the session, so nobody reads it as a verdict on the code.
    """
    module = _load()
    code, output = _run_main(module, _empty_session(tmp_path), tmp_path)

    assert code == 1
    assert "no mutation jobs" in output
    assert "zero mutants were measured" in output
    assert "0/0" not in output


# ---------------------------------------------------------------------------------------------
# The baseline records when each survivor was acquired (#809 T2c)
# ---------------------------------------------------------------------------------------------
#
# ``target_date`` is renegotiable: ``--emit-baseline`` recomputes it as today + horizon, so it
# answers "when was this baseline last written". The acquisition date is the one that must not
# move: it is when the debt was measured, and handing it another 90 days is how a year-old
# ratchet entry stays alive forever.

TODAY = date(2026, 8, 12)


def _entry(module, **overrides):
    fields = {
        "survivors": 1,
        "target": 0,
        "target_date": "2026-08-26",
        "acquired_on": "2026-01-05",
    }
    fields.update(overrides)
    return module.BaselineEntry(**fields)


def test_reemitting_a_baseline_does_not_move_the_acquisition_date() -> None:
    module = _load()
    identity = "app/x.py|operator-0|0"

    payload = module.render_baseline({identity: "app/x.py"}, TODAY, {identity: _entry(module)})

    assert 'acquired_on="2026-01-05"' in payload


def test_a_new_identity_records_the_emission_date() -> None:
    module = _load()

    payload = module.render_baseline({"app/x.py|operator-9|0": "app/x.py"}, TODAY, {})

    assert 'acquired_on="2026-08-12"' in payload


def test_the_horizon_is_still_renegotiated_on_every_emission() -> None:
    """The date that moves is ``target_date``; keeping the old one would be the opposite fix."""
    module = _load()
    identity = "app/x.py|operator-0|0"

    payload = module.render_baseline({identity: "app/x.py"}, TODAY, {identity: _entry(module)})

    assert 'target_date="2026-11-10"' in payload


def test_an_entry_without_an_acquisition_date_starts_its_clock_today() -> None:
    """D7 (1): a baseline emitted before this field must keep loading, and the assumed date is
    written out explicitly so the next reader can see the clock started at the upgrade."""
    module = _load()
    identity = "app/x.py|operator-0|0"

    payload = module.render_baseline(
        {identity: "app/x.py"}, TODAY, {identity: _entry(module, acquired_on="")}
    )

    assert 'acquired_on="2026-08-12"' in payload


def test_an_entry_without_an_acquisition_date_is_asked_to_be_re_emitted() -> None:
    """Without a note, option D7 (1) would be silent forever."""
    module = _load()
    identity = "app/x.py|operator-0|0"
    module.BASELINE = {identity: _entry(module, acquired_on="")}

    violations, notices = module.check_ratchet({identity: "app/x.py"}, TODAY)

    assert violations == []
    assert any("acquired_on" in notice for notice in notices), notices
