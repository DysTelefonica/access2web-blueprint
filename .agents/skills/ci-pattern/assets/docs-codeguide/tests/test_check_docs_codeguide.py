#!/usr/bin/env python3
# test_check_docs_codeguide.py — cobertura del gate docs-codeguide
# `assets/docs-codeguide/check_docs_codeguide.py` (issue #528, T10, A2).
"""Falla cerrada en los dos sentidos del codeguide.

Cobertura de las cuatro reglas del asset:

- Adelante (docs -> repo): toda ruta citada en ``docs/codebase/**``
  existe en el repo (fichero o directorio).
- Atrás (repo -> docs): todo módulo de primer nivel bajo
  ``--code-root`` que case con ``--code-glob`` aparece mencionado en
  alguna página de ``--docs-path``.
- Fail-closed: la entrada vacía en cualquiera de los dos lados es
  FAIL — un codeguide sin páginas o un code-root sin módulos nunca
  declara OK (el comentario de la issue lo prohíbe: «mejor que el
  espejo gentle-ai»).
- Binario: el veredicto es ``exit 0/1/2`` y la causa raíz sale en
  stderr.

El test importa el módulo por ruta, igual que el resto de assets del
repo: ``importlib.util.spec_from_file_location`` contra el script
hermano. Si el script no existe, el ``assert spec and spec.loader``
falla y este test no se puede ejecutar en verde. El estado RED
previo a la implementación es, por tanto, la ausencia del fichero
hermano en el árbol.

Cada test ejerce un camino: el path feliz para asegurar que el
veredicto no se atraganta, y el path de fallo para asegurar que el
modo fail-closed no se relaja. Un test que solo cubra el path verde
sería el agujero exacto que este gate existe para cerrar.
"""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_docs_codeguide.py"
spec = importlib.util.spec_from_file_location("check_docs_codeguide", GATE)
assert spec and spec.loader, (
    f"el gate {GATE.name} no existe todavía; el estado RED previo a la "
    "implementación es la ausencia del fichero hermano"
)
gate = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gate)


# --------------------------------------------------------------------- #
# Helpers de fixture
# --------------------------------------------------------------------- #


def _write(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def _make_repo(root: Path) -> Path:
    """Crea un repo mínimo con un code-root y un docs-path coherentes.

    Estructura::

        repo/
          docs/codebase/mental-model.md
            <- cita `internal/components/app/main.go`
               y `internal/components/billing/core.go`
          internal/components/app/main.go
          internal/components/billing/core.go

    Los módulos viven bajo ``internal/components/`` (el code-root por
    defecto) y ambos nombres aparecen en el doc: el veredicto del
    sentido atrás (repo -> docs) los declara cubiertos.
    """
    repo = root / "repo"
    repo.mkdir()
    _write(repo / "docs" / "codebase" / "mental-model.md",
           "# Mental model\n\n"
           "Usa `internal/components/app/main.go` y "
           "`internal/components/billing/core.go` como entradas.\n")
    _write(repo / "internal" / "components" / "app" / "main.go", "// entry\n")
    _write(repo / "internal" / "components" / "billing" / "core.go", "// billing\n")
    return repo


def _add_codeguide_page(repo: Path, page: str, content: str) -> None:
    _write(repo / "docs" / "codebase" / page, content)


def _add_missing_path(repo: Path, rel: str) -> None:
    """Añade al codeguide una página que cita una ruta inexistente."""
    _add_codeguide_page(repo, "broken.md", f"# Broken\n\nCita `{rel}` que no existe.\n")


def _add_orphan_module(repo: Path, name: str) -> Path:
    """Crea un módulo bajo ``internal/components/<name>/`` que NO aparece
    en el codeguide (huérfano sentido atrás).
    """
    module_dir = repo / "internal" / "components" / name
    module_dir.mkdir(parents=True, exist_ok=True)
    _write(module_dir / "core.go", "// orphan module\n")
    return module_dir


# --------------------------------------------------------------------- #
# Cobertura: extracción de rutas citadas en el codeguide
# --------------------------------------------------------------------- #


class TestExtractPathRefs(unittest.TestCase):
    """El parser de citas: una página Markdown, una lista de rutas."""

    def test_empty_content_returns_empty_list(self) -> None:
        """Página vacía: el parser no inventa rutas."""
        self.assertEqual(gate.extract_path_refs(""), [])

    def test_backtick_quoted_path_is_captured(self) -> None:
        """`` `internal/foo/bar.go` `` aparece como ``internal/foo/bar.go``."""
        refs = gate.extract_path_refs("Usa `internal/foo/bar.go` como entrada.")
        self.assertEqual(refs, ["internal/foo/bar.go"])

    def test_table_cell_with_backtick_path_is_captured(self) -> None:
        """Una fila de tabla `` | `internal/foo/` | owns | `` se extrae."""
        content = (
            "| Path | Owns |\n"
            "|---|---|\n"
            "| `internal/foo/` | X |\n"
        )
        refs = gate.extract_path_refs(content)
        self.assertEqual(refs, ["internal/foo/"])

    def test_non_path_backtick_string_is_ignored(self) -> None:
        """`` `foo` `` sin barra no es una ruta y se ignora."""
        refs = gate.extract_path_refs("El comando `foo` no es una ruta.")
        self.assertEqual(refs, [])

    def test_paths_are_unique_within_a_page(self) -> None:
        """Una página que cita la misma ruta dos veces devuelve una sola entrada."""
        content = (
            "Ver `internal/app/main.go` y, además, `internal/app/main.go`.\n"
        )
        refs = gate.extract_path_refs(content)
        self.assertEqual(refs, ["internal/app/main.go"])

    def test_paths_inside_other_inline_tokens_are_not_misread(self) -> None:
        """Una URL `` `https://example.com/foo` `` no es una ruta de repo."""
        refs = gate.extract_path_refs(
            "Enlace externo `https://example.com/foo` no se trata como ruta."
        )
        self.assertEqual(refs, [])


# --------------------------------------------------------------------- #
# Cobertura: veredicto del gate, sentido adelante (docs -> repo)
# --------------------------------------------------------------------- #


class TestForwardDocsToRepo(unittest.TestCase):
    """Cada ruta citada en ``docs/codebase/**`` existe en el repo."""

    def test_clean_codeguide_passes(self) -> None:
        """Todas las rutas citadas existen: el veredicto es OK (lista vacía)."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            result = gate.check_docs_codeguide(repo)
            self.assertEqual(result["missing_in_repo"], [])
            self.assertEqual(result["missing_in_codeguide"], [])

    def test_missing_path_in_repo_fails_closed(self) -> None:
        """Una página cita ``internal/ghost/foo.go`` que no existe: el gate la nombra."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            _add_codeguide_page(
                repo, "repository-map.md",
                "# Repo map\n\n| Path | Owns |\n|---|---|\n"
                "| `internal/ghost/foo.go` | Inventado |\n",
            )
            result = gate.check_docs_codeguide(repo)
            self.assertIn("internal/ghost/foo.go", result["missing_in_repo"])

    def test_multiple_missing_paths_all_reported(self) -> None:
        """Varias rutas inexistentes: el veredicto lista todas."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            _add_codeguide_page(
                repo, "repository-map.md",
                "| `internal/ghost-a/` | x |\n"
                "| `internal/ghost-b/` | x |\n",
            )
            result = gate.check_docs_codeguide(repo)
            missing = result["missing_in_repo"]
            self.assertIn("internal/ghost-a/", missing)
            self.assertIn("internal/ghost-b/", missing)


# --------------------------------------------------------------------- #
# Cobertura: veredicto del gate, sentido atrás (repo -> docs)
# --------------------------------------------------------------------- #


class TestBackwardRepoToDocs(unittest.TestCase):
    """Cada módulo del code-root aparece mencionado en el codeguide."""

    def test_module_listed_in_codeguide_passes(self) -> None:
        """Un módulo cuyo nombre aparece en alguna página: OK."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            # ``app`` aparece en mental-model.md gracias a la cita
            # `` `internal/components/app/main.go` ``; ``billing``,
            # por la cita `` `internal/components/billing/core.go` ``.
            # Ambos están cubiertos.
            result = gate.check_docs_codeguide(repo)
            self.assertEqual(result["missing_in_codeguide"], [])

    def test_orphan_module_fails_closed(self) -> None:
        """Un módulo nuevo bajo internal/ que nadie cita: el gate lo nombra."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            _add_orphan_module(repo, "ghost")
            result = gate.check_docs_codeguide(repo)
            self.assertIn("ghost", result["missing_in_codeguide"])


# --------------------------------------------------------------------- #
# Cobertura: fail-closed (entrada vacía = fallo, nunca verde)
# --------------------------------------------------------------------- #


class TestFailClosed(unittest.TestCase):
    """La entrada vacía en cualquier lado es un FAIL."""

    def test_empty_docs_codebase_fails(self) -> None:
        """``docs/codebase/`` sin páginas: el gate no dice OK."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "repo"
            repo.mkdir()
            # Sin directorio docs/codebase siquiera.
            result = gate.check_docs_codeguide(repo)
            self.assertEqual(result["missing_in_repo"], ["<empty docs/codebase/>"])

    def test_empty_code_root_fails(self) -> None:
        """Code-root sin módulos: el gate lo nombra como ausente."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "repo"
            repo.mkdir()
            _write(repo / "docs" / "codebase" / "page.md", "# Page\n")
            # Sin internal/components/* (code-root por defecto).
            result = gate.check_docs_codeguide(
                repo, code_root="internal/components", code_glob="*.go",
            )
            self.assertEqual(result["missing_in_codeguide"], ["<empty internal/components/>"])

    def test_empty_repo_path_returns_invalid(self) -> None:
        """El path no es un directorio: la API lanza el error explícito."""
        with self.assertRaises(FileNotFoundError):
            gate.check_docs_codeguide(Path("/path/that/does/not/exist/xyzzy"))


# --------------------------------------------------------------------- #
# Cobertura: configuración por CLI (--docs-path, --code-root, --code-glob)
# --------------------------------------------------------------------- #


class TestConfigurableRoots(unittest.TestCase):
    """Las raíces son configurables; los defaults no se queman en el código."""

    def test_custom_docs_path_is_read(self) -> None:
        """Si el repo usa ``docs/codeguide/`` en lugar de ``docs/codebase/``,
        ``--docs-path`` lo respeta."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "repo"
            repo.mkdir()
            _write(repo / "docs" / "codeguide" / "page.md",
                   "# Page\n\nCita `internal/app/main.go`.\n")
            _write(repo / "internal" / "app" / "main.go", "// entry\n")
            result = gate.check_docs_codeguide(repo, docs_path="docs/codeguide")
            self.assertEqual(result["missing_in_repo"], [])

    def test_custom_code_root_and_glob_are_used(self) -> None:
        """Un repo con módulos Python bajo ``src/modules`` y ``*.py``."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "repo"
            repo.mkdir()
            _write(repo / "docs" / "codebase" / "page.md",
                   "# Page\n\nCubre `src/modules/users.py`.\n")
            _write(repo / "src" / "modules" / "users.py", "# users\n")
            _write(repo / "src" / "modules" / "billing.py", "# billing\n")
            # billing no aparece en el codeguide -> missing_in_codeguide.
            result = gate.check_docs_codeguide(
                repo, code_root="src/modules", code_glob="*.py",
            )
            self.assertIn("billing", result["missing_in_codeguide"])
            self.assertNotIn("users", result["missing_in_codeguide"])


# --------------------------------------------------------------------- #
# Cobertura: contrato del binario (exit 0/1/2)
# --------------------------------------------------------------------- #


class TestCliExitCodes(unittest.TestCase):
    """El binario es el contrato con CI: exit 0 limpio, 1 hallazgos, 2 inválido."""

    def test_cli_exits_zero_on_clean_repo(self) -> None:
        """El path verde: exit 0 y ``OK`` en stdout con el número de páginas."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            proc = subprocess.run(
                [sys.executable, str(GATE), str(repo)],
                capture_output=True, text=True, check=False, encoding="utf-8",
            )
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("OK", proc.stdout)

    def test_cli_exits_one_on_missing_path(self) -> None:
        """El path rojo: exit 1 y nombre de la ruta ausente en stderr."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            _add_missing_path(repo, "internal/ghost/foo.go")
            proc = subprocess.run(
                [sys.executable, str(GATE), str(repo)],
                capture_output=True, text=True, check=False, encoding="utf-8",
            )
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("internal/ghost/foo.go", proc.stderr)

    def test_cli_exits_two_on_missing_repo_path(self) -> None:
        """Input inválido (path que no existe): exit 2 y mensaje en stderr."""
        proc = subprocess.run(
            [sys.executable, str(GATE), "/path/that/does/not/exist/xyzzy"],
            capture_output=True, text=True, check=False, encoding="utf-8",
        )
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("not", proc.stderr.lower() + proc.stdout.lower())


# --------------------------------------------------------------------- #
# Cobertura: contrato del resultado (forma del dict, claves, JSON estable)
# --------------------------------------------------------------------- #


class TestResultContract(unittest.TestCase):
    """El dict ``result`` es el contrato con la API programática."""

    def test_result_has_both_keys(self) -> None:
        """``missing_in_repo`` y ``missing_in_codeguide`` siempre presentes."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            result = gate.check_docs_codeguide(repo)
            self.assertIn("missing_in_repo", result)
            self.assertIn("missing_in_codeguide", result)
            self.assertIsInstance(result["missing_in_repo"], list)
            self.assertIsInstance(result["missing_in_codeguide"], list)

    def test_result_is_json_serializable(self) -> None:
        """El dict se serializa a JSON sin transformaciones; CI y reportadores
        externos lo consumen sin sorpresas."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            result = gate.check_docs_codeguide(repo)
            json.dumps(result)  # no lanza


if __name__ == "__main__":
    unittest.main()
