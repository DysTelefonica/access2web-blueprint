"""La página «dónde trabajar» dice dónde tocar (#1201).

Cubre el gate que exige `docs/codebase/where-to-work.md`: la página está, tiene
las filas canónicas (formulario, job de CI, esquema) más la fila local, cada
fila nombra lee/toca/prueba, y cada ruta citada existe en el repo.

Solo stdlib, ``unittest`` y fixtures en ``tempfile``: el gate lee el árbol y la
página, no la historia de git.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from check_where_to_work import (  # noqa: E402
    DEFAULT_PAGE,
    PageUnreadableError,
    check,
    missing_tasks,
    page_path,
)

PAGE = """# Dónde trabajar

Lee la fila de tu tarea y sigue sus tres columnas.

| Tarea | Lee | Toca | Prueba |
|---|---|---|---|
| `form` | `docs/codebase/backend-modules.md` | `app/modules/animals` | `tests/test_animals.py` |
| `ci` | `docs/codebase/ci-cd.md` | `.github/workflows/ci.yml` | `make verify` |
| `schema` | `docs/codebase/migration.md` | `app/core/migration/sql/001.sql` | `tests/test_schema.py` |
| `local` | `docs/CODEBASE-GUIDE.md` | `skills/architecture-guardrails/SKILL.md` | `make lint` |
"""


def _repo(root: Path, files: dict[str, str]) -> Path:
    """Escribe los ficheros del fixture y devuelve la raíz del repo."""
    for name, content in files.items():
        path = root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
    return root


class TestWhereToWork(unittest.TestCase):
    """El gate mide presencia y existencia de rutas; no juzga la prosa."""

    def test_a_complete_page_has_no_findings(self) -> None:
        """C1: la página canónica pasa."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(
                Path(tmp),
                {
                    "docs/codebase/where-to-work.md": PAGE,
                    "docs/codebase/backend-modules.md": "x\n",
                    "docs/codebase/ci-cd.md": "x\n",
                    "docs/codebase/migration.md": "x\n",
                    "docs/CODEBASE-GUIDE.md": "x\n",
                    "app/modules/animals/routes.py": "x = 1\n",
                    ".github/workflows/ci.yml": "name: ci\n",
                    "app/core/migration/sql/001.sql": "select 1;\n",
                    "skills/architecture-guardrails/SKILL.md": "# skill\n",
                    "tests/test_animals.py": "def test_x(): pass\n",
                    "tests/test_schema.py": "def test_x(): pass\n",
                },
            )
            self.assertEqual([], check(repo))

    def test_the_page_is_the_documented_path(self) -> None:
        """C1: la página vive en `docs/codebase/where-to-work.md` por defecto."""
        self.assertEqual(Path("/repo/docs/codebase/where-to-work.md"), page_path(Path("/repo")))
        self.assertEqual("docs/codebase/where-to-work.md", DEFAULT_PAGE)

    def test_a_missing_page_is_a_finding(self) -> None:
        """C1: sin página, el gate dice que falta."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {"README.md": "x\n"})
            findings = check(repo)
            self.assertEqual(1, len(findings))
            self.assertIn("where-to-work.md", findings[0])

    def test_a_missing_canonical_row_is_a_finding(self) -> None:
        """C1: sin la fila de esquema, el gate la nombra."""
        with tempfile.TemporaryDirectory() as tmp:
            page = "\n".join(line for line in PAGE.splitlines() if "`schema`" not in line)
            repo = _repo(Path(tmp), {"docs/codebase/where-to-work.md": page})
            self.assertIn("schema", " ".join(check(repo)))

    def test_the_missing_rows_are_named(self) -> None:
        """C1: el gate enumera las filas canónicas que faltan."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {"docs/codebase/where-to-work.md": "# vacío\n"})
            self.assertEqual({"form", "ci", "schema", "local"}, missing_tasks(repo))

    def test_an_empty_cell_is_a_finding(self) -> None:
        """C1: una fila sin «toca» no sirve; el gate la nombra."""
        with tempfile.TemporaryDirectory() as tmp:
            page = PAGE.replace(
                "| `ci` | `docs/codebase/ci-cd.md` | `.github/workflows/ci.yml` | `make verify` |",
                "| `ci` | `docs/codebase/ci-cd.md` |  | `make verify` |",
            )
            repo = _repo(Path(tmp), {"docs/codebase/where-to-work.md": page})
            self.assertIn("ci", " ".join(check(repo)))

    def test_a_dead_path_is_a_finding(self) -> None:
        """C1/C4: una ruta citada que ya no existe es deriva real."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {"docs/codebase/where-to-work.md": PAGE})
            findings = check(repo)
            self.assertTrue(
                any("app/modules/animals" in finding for finding in findings),
                findings,
            )

    def test_placeholders_are_not_paths(self) -> None:
        """C4: un marcador de posición no es una ruta que pueda existir."""
        with tempfile.TemporaryDirectory() as tmp:
            page = PAGE.replace("`app/modules/animals`", "`app/modules/<módulo>/`")
            repo = _repo(
                Path(tmp),
                {
                    "docs/codebase/where-to-work.md": page,
                    "docs/codebase/backend-modules.md": "x\n",
                    "docs/codebase/ci-cd.md": "x\n",
                    "docs/codebase/migration.md": "x\n",
                    "docs/CODEBASE-GUIDE.md": "x\n",
                    ".github/workflows/ci.yml": "name: ci\n",
                    "app/core/migration/sql/001.sql": "select 1;\n",
                    "skills/architecture-guardrails/SKILL.md": "# skill\n",
                    "tests/test_animals.py": "def test_x(): pass\n",
                    "tests/test_schema.py": "def test_x(): pass\n",
                },
            )
            self.assertEqual([], check(repo))

    def test_a_page_without_the_three_columns_is_a_finding(self) -> None:
        """C1: la tabla declara lee, toca y prueba."""
        with tempfile.TemporaryDirectory() as tmp:
            page = "| Tarea | Dónde |\n|---|---|\n| `form` | `docs/x.md` |\n"
            repo = _repo(Path(tmp), {"docs/codebase/where-to-work.md": page})
            self.assertIn("lee", " ".join(check(repo)))

    def test_an_unreadable_page_fails_closed(self) -> None:
        """C4: una página ilegible es rojo, no verde."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _repo(Path(tmp), {"docs/codebase/where-to-work.md": "# ok\n"})
            (repo / "docs/codebase/where-to-work.md").write_bytes(b"\xff\xfe\x00")
            with self.assertRaises(PageUnreadableError):
                check(repo)


if __name__ == "__main__":
    unittest.main()
