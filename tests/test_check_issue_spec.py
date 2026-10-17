"""Smoke test for ``scripts/check_issue_spec.py`` (canon de DysTelefonica/team-skills#633).

Stub temporal: la suite canónica de 35 tests entra con PR-2 (chain 2/3, Refs
#960). Mientras tanto, este stub mantiene verde el job ``tests`` del CI: el
canon byte a byte cambia la API pública (``validate`` -> ``validate_body`` /
``validate_remote``) y los tests viejos de #855 fallan hasta que entren los
del canon.
"""

from __future__ import annotations

import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import check_issue_spec  # noqa: E402

REPO = "DysTelefonica/access2web-blueprint"


def test_references_extracts_closes_keyword() -> None:
    refs = check_issue_spec.references("Closes #42")
    assert [(r.kind, r.number) for r in refs] == [("close", 42)]


def test_validate_body_empty_reports_no_reference_with_depends_on_hint() -> None:
    violations = check_issue_spec.validate_body("", REPO)
    assert len(violations) == 1
    assert "depends-on" in violations[0]


def test_validate_body_closes_and_refs_same_issue_fails() -> None:
    violations = check_issue_spec.validate_body("Closes #42\nRefs #42\n", REPO)
    assert any("MISMA issue" in v for v in violations)


def test_validate_body_other_repository_reference_fails() -> None:
    violations = check_issue_spec.validate_body("Closes other/repo#42\n", REPO)
    assert any("otro repositorio" in v for v in violations)


def test_cli_local_mode_passes_with_reference(tmp_path: Path) -> None:
    body_file = tmp_path / "body.md"
    body_file.write_text("Refs #42\n", encoding="utf-8")
    rc = check_issue_spec.main(["local", "--body-file", str(body_file), "--repository", REPO])
    assert rc == 0
