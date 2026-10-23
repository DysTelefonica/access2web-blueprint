#!/usr/bin/env python3
# test_check_docs_codeguide.py — cobertura del gate docs-codeguide
# `assets/docs-codeguide/check_docs_codeguide.py` (issue #528, T10, A2).
"""Falla cerrada en los dos sentidos del codeguide.

Cobertura de las cuatro reglas del asset:

- Adelante (docs -> repo): toda ruta citada en ``docs/codebase/**``
  existe en el repo (fichero o directorio).
- Atrás (repo -> docs): todo módulo de primer nivel bajo
  ``--code-root`` que case con ``--code-glob`` está nombrado en alguna
  página de ``--docs-path``: el nombre completo y en contexto de código
  (backticks, bloque cercado) o como segmento de una ruta (#1004). La
  prosa no cuenta, porque «nuevo» y «datos» salían en cualquier mapa.
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


class TestFencedBlocksAreNotCitations(unittest.TestCase):
    """Un bloque cercado es un ejemplo, no una cita de ruta (issue #725).

    El regex de comillas invertidas empareja tres de las seis comillas de
    un bloque y deja dos sueltas; sin saltar los bloques, el código en
    línea posterior queda con el emparejamiento invertido.
    """

    def test_slash_inside_backtick_fence_is_not_a_path(self) -> None:
        """Falso positivo: el cuerpo de un bloque no es una cita."""
        content = "```python\nfrom src/mod import a\n```\n\nVea `src/mod/a.py`.\n"
        self.assertEqual(gate.extract_path_refs(content), ["src/mod/a.py"])

    def test_inline_path_after_fence_is_still_extracted(self) -> None:
        """Falso negativo: la ruta tras un bloque se extrae, no se pierde."""
        content = "```text\nx\n```\n\nVea `src/ghost/gone.py`.\n"
        self.assertEqual(gate.extract_path_refs(content), ["src/ghost/gone.py"])

    def test_broken_path_after_fence_fails_closed(self) -> None:
        """Extremo a extremo: una ruta rota tras un bloque deja el gate en rojo."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            _add_codeguide_page(
                repo, "after.md", "# After\n\n```text\nx\n```\n\nVea `internal/ghost/gone.go`.\n"
            )
            result = gate.check_docs_codeguide(repo)
            self.assertEqual(result["missing_in_repo"], ["internal/ghost/gone.go"])

    def test_inline_paths_around_a_fence_are_both_extracted(self) -> None:
        content = "Antes `a/b.py`.\n\n```\nc/d.py\n```\n\nDespués `e/f.py`.\n"
        self.assertEqual(gate.extract_path_refs(content), ["a/b.py", "e/f.py"])

    def test_tilde_fence_is_ignored(self) -> None:
        content = "~~~bash\ncd src/mod\n~~~\n\nVea `src/mod/a.py`.\n"
        self.assertEqual(gate.extract_path_refs(content), ["src/mod/a.py"])

    def test_longer_fence_contains_a_shorter_one(self) -> None:
        """Un bloque de cuatro comillas solo lo cierra otro de al menos cuatro."""
        content = "````md\n```python\nfrom a/b import c\n```\nd/e.py\n````\n\nVea `x/y.py`.\n"
        self.assertEqual(gate.extract_path_refs(content), ["x/y.py"])

    def test_tilde_fence_is_not_closed_by_backtick_fence(self) -> None:
        content = "~~~\n```\na/b.py\n~~~\n\nVea `x/y.py`.\n"
        self.assertEqual(gate.extract_path_refs(content), ["x/y.py"])

    def test_indented_fence_in_a_list_item_is_ignored(self) -> None:
        content = "1. Paso:\n\n   ```bash\n   cd src/mod\n   ```\n\nVea `x/y.py`.\n"
        self.assertEqual(gate.extract_path_refs(content), ["x/y.py"])

    def test_unclosed_fence_discards_until_end_of_page(self) -> None:
        content = "Vea `x/y.py`.\n\n```text\na/b.py\n`c/d.py`\n"
        self.assertEqual(gate.extract_path_refs(content), ["x/y.py"])


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


class TestGlobsMarkersAndBraces(unittest.TestCase):
    """#737: globs con coincidencias, marcadores con prefijo existente y llaves son citas válidas."""

    def _missing(self, cite: str, extra: dict[str, str] | None = None) -> list[str]:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            for rel, body in (extra or {}).items():
                _write(repo / rel, body)
            _add_codeguide_page(repo, "cites.md", f"Vea `{cite}`.\n")
            return gate.check_docs_codeguide(repo)["missing_in_repo"]

    def test_glob_matching_a_file_is_valid(self) -> None:
        self.assertEqual(self._missing("internal/components/app/*.go"), [])

    def test_glob_matching_nothing_is_red_and_named(self) -> None:
        found = self._missing("internal/components/app/*.rs")
        self.assertEqual(len(found), 1)
        self.assertIn("internal/components/app/*.rs", found[0])
        self.assertIn("glob sin coincidencias", found[0])

    def test_marker_with_existing_prefix_is_valid(self) -> None:
        self.assertEqual(self._missing("internal/components/<module>/core.go"), [])

    def test_marker_with_missing_prefix_is_red_and_named(self) -> None:
        found = self._missing("internal/ghost/<module>/core.go")
        self.assertEqual(len(found), 1)
        self.assertIn("internal/ghost/<module>/core.go", found[0])
        self.assertIn("prefijo inexistente: internal/ghost", found[0])

    def test_glob_before_a_marker_is_checked_as_prefix(self) -> None:
        self.assertEqual(self._missing("internal/*/<module>/core.go"), [])
        self.assertTrue(self._missing("nada/*/<module>/core.go"))

    def test_marker_followed_by_glob_only_checks_the_prefix(self) -> None:
        self.assertEqual(self._missing("internal/components/<module>/*.go"), [])

    def test_literal_missing_path_keeps_its_old_message(self) -> None:
        self.assertEqual(self._missing("internal/ghost/foo.go"), ["internal/ghost/foo.go"])

    def test_braces_are_valid_when_every_expansion_is(self) -> None:
        self.assertEqual(
            self._missing("internal/components/{app/main,billing/core}.go"), [])
        self.assertEqual(self._missing("internal/components/<m>/{main,core}.go"), [])

    def test_braces_fail_when_one_expansion_is_not_valid(self) -> None:
        found = self._missing("internal/components/{app/main,ghost/main}.go")
        self.assertEqual(len(found), 1)
        self.assertIn("internal/components/{app/main,ghost/main}.go", found[0])
        self.assertIn("internal/components/ghost/main.go", found[0])

    def test_glob_char_in_a_directory_name_resolves_by_glob(self) -> None:
        self.assertEqual(self._missing("internal/comp*/app/main.go"), [])

    def test_cli_exits_one_on_glob_without_matches(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = _make_repo(Path(tmp))
            _add_codeguide_page(repo, "cites.md", "Vea `internal/components/app/*.rs`.\n")
            result = subprocess.run(
                [sys.executable, str(GATE), str(repo)],
                capture_output=True, text=True, check=False)
            self.assertEqual(result.returncode, 1)
            self.assertIn("glob sin coincidencias", result.stderr)


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


class TestMultipleRoots(unittest.TestCase):
    """#727: varias raíces de código, cada una con su glob."""

    @staticmethod
    def _repo_with_two_roots(root: Path, page: str) -> Path:
        repo = root / "repo"
        repo.mkdir()
        _write(repo / "docs" / "codebase" / "page.md", page)
        _write(repo / "app" / "users.py", "# users\n")
        _write(repo / "templates" / "home.html", "<p>hola</p>\n")
        return repo

    def test_every_root_covered_passes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo_with_two_roots(Path(tmp), "`app/users.py` y `templates/home.html`\n")
            result = gate.check_docs_codeguide_roots(
                repo, [("app", "*.py"), ("templates", "*.html")]
            )
            self.assertEqual(result, {"missing_in_repo": [], "missing_in_codeguide": []})

    def test_module_missing_in_second_root_names_root_and_module(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo_with_two_roots(Path(tmp), "`app/users.py`\n")
            result = gate.check_docs_codeguide_roots(
                repo, [("app", "*.py"), ("templates", "*.html")]
            )
            self.assertEqual(result["missing_in_codeguide"], ["home (templates)"])

    def test_empty_second_root_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo_with_two_roots(Path(tmp), "`app/users.py` y home\n")
            result = gate.check_docs_codeguide_roots(
                repo, [("app", "*.py"), ("static", "*.js")]
            )
            self.assertEqual(result["missing_in_codeguide"], ["<empty static/>"])

    def test_cited_missing_path_is_reported_once_across_roots(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo_with_two_roots(Path(tmp), "`app/users.py`, home y `app/ghost.py`\n")
            result = gate.check_docs_codeguide_roots(
                repo, [("app", "*.py"), ("templates", "*.html")]
            )
            self.assertEqual(result["missing_in_repo"], ["app/ghost.py"])

    def test_single_root_keeps_bare_module_names(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo_with_two_roots(Path(tmp), "nada\n")
            result = gate.check_docs_codeguide_roots(repo, [("app", "*.py")])
            self.assertEqual(result["missing_in_codeguide"], ["users"])

    def test_cli_repeated_pairs_exit_zero_when_all_roots_are_covered(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo_with_two_roots(Path(tmp), "`app/users.py` y `templates/home.html`\n")
            proc = subprocess.run(
                [sys.executable, str(GATE), str(repo), "--code-root", "app",
                 "--code-glob", "*.py", "--code-root", "templates", "--code-glob", "*.html"],
                capture_output=True, text=True, check=False, encoding="utf-8",
            )
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_cli_repeated_pairs_exit_one_when_a_root_has_an_undescribed_module(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo_with_two_roots(Path(tmp), "`app/users.py`\n")
            proc = subprocess.run(
                [sys.executable, str(GATE), str(repo), "--code-root", "app",
                 "--code-glob", "*.py", "--code-root", "templates", "--code-glob", "*.html"],
                capture_output=True, text=True, check=False, encoding="utf-8",
            )
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("home (templates)", proc.stderr)

    def test_cli_unpaired_roots_and_globs_exit_two(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo_with_two_roots(Path(tmp), "`app/users.py`\n")
            proc = subprocess.run(
                [sys.executable, str(GATE), str(repo), "--code-root", "app",
                 "--code-root", "templates", "--code-glob", "*.py"],
                capture_output=True, text=True, check=False, encoding="utf-8",
            )
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("code-root", proc.stderr)


class TestGlobMatchesEquivalence(unittest.TestCase):
    """#757: ``_glob_matches`` usa ``Path.glob``; el comportamiento observable se mantiene."""

    def _repo(self, tmp: str) -> Path:
        repo = Path(tmp)
        _write(repo / "src" / "a" / "deep" / "m.py", "x = 1\n")
        _write(repo / "src" / "b.py", "y = 1\n")
        _write(repo / "src" / ".hidden.py", "z = 1\n")
        return repo

    def test_star_and_recursive_patterns_match(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo(tmp)
            self.assertTrue(gate._glob_matches(repo, "src/*.py"))
            self.assertTrue(gate._glob_matches(repo, "src/**/*.py"))
            self.assertTrue(gate._glob_matches(repo, "src/a/**/m.py"))

    def test_pattern_without_matches_does_not_match(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo(tmp)
            self.assertFalse(gate._glob_matches(repo, "src/*.go"))
            self.assertFalse(gate._glob_matches(repo, "nothing/**/*.py"))

    def test_hidden_file_counts_as_a_match(self) -> None:
        """Único cambio respecto a ``glob.glob``: los ficheros ocultos también casan."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            _write(repo / "src" / ".hidden.py", "z = 1\n")
            self.assertTrue(gate._glob_matches(repo, "src/*.py"))

    def test_unsupported_patterns_do_not_match_and_do_not_raise(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._repo(tmp)
            self.assertFalse(gate._glob_matches(repo, "/etc/*"))
            self.assertFalse(gate._glob_matches(repo, ""))

    def test_repo_path_with_glob_metacharacters_is_literal(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "we[ir]d*dir"
            _write(repo / "src" / "b.py", "y = 1\n")
            self.assertTrue(gate._glob_matches(repo, "src/*.py"))


class TestPublicErrorTypes(unittest.TestCase):
    """#757: los errores con mensaje propio siguen siendo ``ValueError`` y ``FileNotFoundError``."""

    def test_no_roots_is_a_value_error_with_the_same_message(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(ValueError) as ctx:
                gate.check_docs_codeguide_roots(Path(tmp), [])
            self.assertEqual(str(ctx.exception), "at least one (code_root, code_glob) pair is required")

    def test_missing_repo_is_a_file_not_found_error_with_the_same_message(self) -> None:
        with self.assertRaises(FileNotFoundError) as ctx:
            gate.check_docs_codeguide_roots(Path("/path/that/does/not/exist"), [("a", "*.py")])
        self.assertEqual(str(ctx.exception), "repo path is not a directory: /path/that/does/not/exist")


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


class TestModuleNameIsNotAProseSubstring(unittest.TestCase):
    """team-skills#1004: un nombre de módulo cuenta cuando el mapa lo nombra
    como código, no cuando aparece por casualidad en la prosa.

    El sentido atrás comparaba el nombre con ``in`` sobre el texto entero
    del mapa. Eso daba por mapeado un módulo sin mapear en cuanto su
    nombre era una palabra corriente del español (``nuevo``, ``datos``) o
    un trozo de otra palabra (``core`` dentro de ``encore``). Medido en
    BP: ``app/src/modules/nuevo/`` pasaba el gate porque «nuevo» salía 15
    veces como adjetivo.

    Cada test de este bloque monta el caso completo — una página del mapa
    que menciona el nombre en prosa y un módulo en el árbol que el gate
    tiene que nombrar— y comprueba el veredicto a través de la función
    pública, no de un ayudante interno.
    """

    def _repo_with_module_and_page(self, root: Path, module: str, page: str) -> Path:
        repo = root / "repo"
        _write(repo / "docs" / "codebase" / "page.md", page)
        _write(repo / "src" / "modules" / module / "__init__.py", "# " + module + "\n")
        return repo

    def _verdict(self, tmp: str, module: str, page: str) -> list[str]:
        repo = self._repo_with_module_and_page(Path(tmp), module, page)
        result = gate.check_docs_codeguide(
            repo, docs_path="docs/codebase", code_root="src/modules", code_glob="*.py",
        )
        return result["missing_in_codeguide"]

    def test_module_named_only_as_a_spanish_adjective_is_missing(self) -> None:
        """``nuevo``: el mapa dice «nuevos módulos», «nueva regla», nunca el
        nombre del módulo. Es el caso medido en BP."""
        page = (
            "# Guía\n\n"
            "El árbol tiene nuevos módulos y cada uno recibe una nueva regla.\n\n"
            "Los nuevos cambios llegan por PR; la nueva plantilla vive en el repo.\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(self._verdict(tmp, "nuevo", page), ["nuevo"])

    def test_module_named_only_in_prose_and_inside_another_word_is_missing(self) -> None:
        """``datos``: «los datos del usuario» es una palabra entera y aun así
        no nombra un módulo; ``bigdatos`` es además un trozo de otra palabra."""
        page = (
            "# Guía\n\n"
            "Los datos del usuario viven en su perfil y los bigdatos se archivan.\n\n"
            "Este repositorio guarda datos y exporta datos a diario.\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(self._verdict(tmp, "datos", page), ["datos"])

    def test_module_named_only_as_a_fragment_of_another_word_is_missing(self) -> None:
        """``core``: no es prosa, es un fragmento: ``encore``, ``scores`` y el
        acento de «núcleo» no son el nombre del módulo."""
        page = (
            "# Guía\n\n"
            "La sección encore llega después, y los scores se publican aparte.\n\n"
            "El núcleo del sistema se documenta en su propia página.\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(self._verdict(tmp, "core", page), ["core"])

    def test_module_named_in_backticks_is_mapped(self) -> None:
        """Camino verde: `` `lanzadera` `` en backticks sí nombra el módulo."""
        page = (
            "# Guía\n\n"
            "Los datos del usuario y los nuevos módulos no dicen nada del código.\n\n"
            "El módulo `lanzadera` es la madre; `nuevo` es un adjetivo.\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(self._verdict(tmp, "lanzadera", page), [])

    def test_module_named_as_a_path_segment_is_mapped(self) -> None:
        """Camino verde: el nombre como segmento de ruta también lo nombra,
        sin backticks. Es como lo cita un mapa que dibuja el árbol."""
        page = (
            "# Guía\n\n"
            "El árbol cuelga de app/src/modules/expedientes/dominio.py y sus\n"
            "vecinos; los datos del usuario nomapped aquí.\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(self._verdict(tmp, "expedientes", page), [])

    def test_module_named_in_a_fenced_block_is_mapped(self) -> None:
        """Camino verde: un bloque cercado también es contexto de código."""
        page = (
            "# Guía\n\n"
            "Los datos van fuera de código.\n\n"
            "```\napp/src/modules/nuevo/\n```\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(self._verdict(tmp, "nuevo", page), [])

    def test_module_named_as_a_file_with_extension_in_backticks_is_mapped(self) -> None:
        """Caso real de APAP_WEB: el módulo ``index`` de
        ``app/templates/index.html`` solo aparece como `` `index.html` ``.
        El punto es un límite, no un carácter del nombre, así que el
        nombre completo sigue contando."""
        page = (
            "# Guía\n\n"
            "Los datos del usuario no dicen nada del código.\n\n"
            "La portada vive en `index.html` y se llama index.\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(self._verdict(tmp, "index", page), [])


if __name__ == "__main__":
    unittest.main()
