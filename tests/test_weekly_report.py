"""Status of each migration milestone (issue #1802).

``scripts/weekly_report.py`` reads the milestones (``hitos``) of
``docs/07-migracion/seguimiento/fases.json`` and the issues labelled
``saneamiento:coexistencia``, and prints one line per milestone. A milestone is
closed only in its estimated week and a delay is never hidden. These tests feed
the input with ``--from-json``: no network.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "scripts" / "weekly_report.py"
PHASES = ROOT / "docs" / "07-migracion" / "seguimiento" / "fases.json"


def _issue(
    repo: str,
    number: int,
    state: str,
    title: str,
    body: str = "",
    labels: tuple[str, ...] = (),
) -> dict[str, object]:
    return {
        "repository": {"name": repo},
        "number": number,
        "state": state,
        "title": title,
        "url": f"https://github.com/DysTelefonica/{repo}/issues/{number}",
        "body": body,
        "createdAt": "2026-11-20T09:00:00Z",
        "labels": [{"name": label} for label in labels],
    }


def _phases(
    tmp_path: Path,
    weeks: dict[str, str],
    replans: dict[str, list[dict[str, str]]] | None = None,
) -> Path:
    definition = json.loads(PHASES.read_text(encoding="utf-8"))
    for hito in definition["hitos"]:
        hito["semana_estimada"] = weeks.get(hito["id"])
        hito["replanificaciones"] = (replans or {}).get(hito["id"], [])
        hito["tasks"] = []
    path = tmp_path / "fases.json"
    path.write_text(json.dumps(definition), encoding="utf-8")
    return path


def _run(
    tmp_path: Path, data: dict[str, object], date: str, phases: Path = PHASES
) -> subprocess.CompletedProcess[str]:
    source = tmp_path / f"input-{date}.json"
    source.write_text(json.dumps(data), encoding="utf-8")
    out_dir = tmp_path / "out"
    out_dir.mkdir(exist_ok=True)
    command = [sys.executable, str(SCRIPT), "--from-json", str(source), "--date", date]
    command += ["--phases", str(phases), "--out-dir", str(out_dir)]
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    return subprocess.run(
        command, capture_output=True, text=True, encoding="utf-8", check=False, env=env
    )


IDCORREO = "refactor: dejar de calcular IDCorreo al encolar correos"
PROBLEM = "### Problema y contexto\n\n{}\n\n### Evidencia verificable\n\nOtra cosa."
CLOSED_DECEMBER = "semana del 21 de diciembre de 2026"


def test_the_real_milestones_count_issues_and_their_own_tasks(tmp_path: Path) -> None:
    result = _run(tmp_path, {"issues": [_issue("BRASS", 7, "open", IDCORREO)]}, "2026-11-25")
    assert result.returncode == 0, result.stderr
    # The real fases.json: one issue plus the static task of the hito, and its week.
    assert f"Hito 2026-12: 0/2 (0 %), {CLOSED_DECEMBER}, en preparación" in result.stdout
    assert "Hito 2027-02: 0/5 (0 %), semana por definir, en preparación" in result.stdout


def test_a_hito_without_estimated_week_shows_it_undefined_and_never_closes(
    tmp_path: Path,
) -> None:
    issues = [_issue("BRASS", 7, "closed", IDCORREO), _issue("HPS", 3, "closed", IDCORREO)]
    result = _run(tmp_path, {"issues": issues}, "2027-01-15", _phases(tmp_path, {}))
    assert result.returncode == 0, result.stderr
    assert "Hito 2026-12: 2/2 (100 %), semana por definir, en plazo" in result.stdout


@pytest.mark.parametrize(
    ("date", "states", "line"),
    [
        (
            "2026-12-28",
            ("closed", "closed", "open"),
            "Hito 2027-01: 2/3 (67 %), semana del 25 de enero de 2027, en preparación",
        ),
        (
            "2027-01-20",
            ("closed", "closed", "closed"),
            "Hito 2027-01: 3/3 (100 %), semana del 25 de enero de 2027, en plazo",
        ),
    ],
)
def test_the_current_hito_shows_real_progress_and_stays_in_plazo_before_its_week(
    tmp_path: Path, date: str, states: tuple[str, str, str], line: str
) -> None:
    phases = _phases(tmp_path, {"2026-12": "2026-12-21", "2027-01": "2027-01-25"})
    issues = [
        _issue("BRASS", 7, "closed", IDCORREO),
        _issue("Lanzadera", 39, states[0], "feat: migraciones"),
        _issue("Lanzadera", 40, states[1], "fix: clave"),
        _issue("Expedientes", 5, states[2], "fix: fecha"),
    ]
    result = _run(tmp_path, {"issues": issues}, date, phases)
    assert result.returncode == 0, result.stderr
    assert line in result.stdout.splitlines()


def test_all_done_in_its_week_closes_the_hito_and_the_next_becomes_current(
    tmp_path: Path,
) -> None:
    phases = _phases(tmp_path, {"2026-12": "2026-12-21"})
    issues = [_issue("BRASS", 7, "closed", IDCORREO), _issue("Lanzadera", 40, "open", "fix")]
    before = _run(tmp_path, {"issues": issues}, "2026-12-20", phases)
    assert f"Hito 2026-12: 1/1 (100 %), {CLOSED_DECEMBER}, en plazo" in before.stdout
    result = _run(tmp_path, {"issues": issues}, "2026-12-21", phases)
    assert f"Hito 2026-12: 1/1 (100 %), {CLOSED_DECEMBER}, cerrado" in result.stdout


@pytest.mark.parametrize(
    "replans",
    [[], [{"fecha": "2026-12-20", "semana_estimada": "2027-01-11", "motivo": "Baja."}]],
)
def test_a_hito_past_its_week_with_pending_work_is_delayed_even_after_a_replan(
    tmp_path: Path, replans: list[dict[str, str]]
) -> None:
    phases = _phases(tmp_path, {"2026-12": "2026-12-14"}, {"2026-12": replans})
    issues = [_issue("BRASS", 7, "open", IDCORREO)]
    result = _run(tmp_path, {"issues": issues}, "2026-12-28", phases)
    assert result.returncode == 0, result.stderr
    assert result.stdout.startswith("Hito 2026-12: 0/1 (0 %)")
    assert ", retrasado" in result.stdout.splitlines()[0]


def test_the_snapshot_has_one_line_per_record_and_no_issue_bodies(tmp_path: Path) -> None:
    closed = _issue("BRASS", 7, "closed", IDCORREO, PROBLEM.format("Correo duplicado. Más."))
    closed["closedAt"] = "2026-11-24T10:00:00Z"
    issues = [closed, *(_issue("HPS", n, "open", "fix", PROBLEM.format("x.")) for n in range(30))]
    # gh search reports the zero time as closedAt of an open issue.
    issues[1]["closedAt"] = "0001-01-01T00:00:00Z"
    assert _run(tmp_path, {"issues": issues}, "2026-11-25").returncode == 0
    text = (tmp_path / "out" / "semanal" / "2026-11-25.json").read_text(encoding="utf-8")
    snapshot = json.loads(text)
    records = sum(len(v) for v in snapshot.values() if isinstance(v, list))
    assert len(text.splitlines()) <= records + 15
    assert '"body"' not in text and "Problema y contexto" not in text
    assert snapshot["issues"][0] == {
        "key": "BRASS#7",
        "state": "closed",
        "hito": "2026-12",
        "title": IDCORREO,
        "url": "https://github.com/DysTelefonica/BRASS/issues/7",
        "createdAt": "2026-11-20T09:00:00Z",
        "closedAt": "2026-11-24T10:00:00Z",
        "labels": [],
        "motivo": "Correo duplicado.",
    }
    assert snapshot["issues"][1]["closedAt"] is None


def test_a_phase_one_group_in_no_hito_fails_naming_it_and_writes_nothing(tmp_path: Path) -> None:
    definition = json.loads(PHASES.read_text(encoding="utf-8"))
    definition["hitos"][1]["groups"].remove("Condor")
    phases = tmp_path / "fases.json"
    phases.write_text(json.dumps(definition), encoding="utf-8")
    result = _run(tmp_path, {"issues": []}, "2026-11-25", phases)
    assert result.returncode == 2
    assert "Condor" in result.stderr
    assert list((tmp_path / "out").iterdir()) == []
