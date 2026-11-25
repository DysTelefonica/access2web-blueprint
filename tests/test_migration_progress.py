"""Migration progress report generated from GitHub issues (issue #1794).

The report is never edited by hand: ``scripts/migration_progress.py`` reads the
phase definition (``docs/07-migracion/seguimiento/fases.json``) and the issues
labelled ``saneamiento:coexistencia``, and writes ``avance.md``. These tests
feed the issues with ``--from-json`` so no network call is made.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "scripts" / "migration_progress.py"


def _issue(repo: str, number: int, state: str, title: str) -> dict[str, object]:
    return {
        "repository": {"name": repo, "nameWithOwner": f"DysTelefonica/{repo}"},
        "number": number,
        "state": state,
        "title": title,
        "url": f"https://github.com/DysTelefonica/{repo}/issues/{number}",
    }


def _run(tmp_path: Path, issues: list[dict[str, object]]) -> subprocess.CompletedProcess[str]:
    source = tmp_path / "issues.json"
    source.write_text(json.dumps(issues), encoding="utf-8")
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    return subprocess.run(
        [
            sys.executable,
            str(SCRIPT),
            "--from-json",
            str(source),
            "--date",
            "2026-11-25",
            "--out-dir",
            str(out_dir),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )


def _markdown(tmp_path: Path) -> str:
    return (tmp_path / "out" / "avance.md").read_text(encoding="utf-8")


SAMPLE = [
    _issue("Lanzadera", 39, "closed", "feat: módulo de migraciones"),
    _issue("Lanzadera", 40, "open", "fix: clave primaria"),
    _issue("Expedientes", 5, "open", "fix: tipo de fecha"),
]


def test_a_closed_issue_counts_as_done_and_groups_carry_their_totals(tmp_path: Path) -> None:
    result = _run(tmp_path, SAMPLE)
    assert result.returncode == 0, result.stderr
    markdown = _markdown(tmp_path)
    assert "### Lanzadera (1/2)" in markdown
    assert "### Expedientes (0/1)" in markdown
    assert (
        "- [x] [Lanzadera#39](https://github.com/DysTelefonica/Lanzadera/issues/39) "
        "feat: módulo de migraciones"
    ) in markdown
    assert "- [ ] [Lanzadera#40]" in markdown
    assert markdown.index("### Lanzadera") < markdown.index("### Expedientes")


def test_an_idcorreo_issue_goes_to_block_a_whatever_its_repo(tmp_path: Path) -> None:
    issues = [
        _issue("BRASS", 7, "open", "refactor: dejar de calcular IDCorreo al encolar correos"),
        _issue("access2web-blueprint", 1741, "closed", "refactor(platform): IDCorreo autonumérico"),
        _issue("BRASS", 8, "open", "fix: otra cosa"),
    ]
    result = _run(tmp_path, issues)
    assert result.returncode == 0, result.stderr
    markdown = _markdown(tmp_path)
    block_a = markdown.split("### Bloque A. IDCorreo (1/2)", 1)[1].split("###", 1)[0]
    assert "[BRASS#7]" in block_a
    assert "[access2web-blueprint#1741]" in block_a
    assert "[BRASS#8]" not in block_a
    assert "### Brass (0/1)" in markdown
    assert markdown.index("### Bloque A. IDCorreo") < markdown.index("### Lanzadera")


def test_an_issue_from_an_unmapped_repo_fails_naming_it_and_writes_nothing(
    tmp_path: Path,
) -> None:
    result = _run(tmp_path, [*SAMPLE, _issue("SICA", 3, "open", "fix: algo")])
    assert result.returncode == 2
    assert "SICA" in result.stderr
    assert list((tmp_path / "out").iterdir()) == []


def test_the_summary_row_of_phase_one_adds_static_tasks_and_issues(tmp_path: Path) -> None:
    result = _run(tmp_path, SAMPLE)
    assert result.returncode == 0, result.stderr
    markdown = _markdown(tmp_path)
    assert "| 0 | Inventario | revisada | 4 | 4 |" in markdown
    assert "| 1 | Saneamiento en Access | en curso | 1 | 6 |" in markdown
    assert "No edite este fichero a mano" in markdown
    assert "2026-11-25" in markdown
