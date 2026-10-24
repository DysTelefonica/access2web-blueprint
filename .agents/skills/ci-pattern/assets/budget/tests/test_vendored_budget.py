#!/usr/bin/env python3
"""Suite del módulo compartido del presupuesto de revisión (#931).

Cubre las reglas de ``assets/budget/vendored_budget.py`` directamente. Las
pruebas de paridad con los gates ``pr-size`` y ``pr-contract`` viven en
``test_budget_parity.py``, que se añade con el cableado de los gates. El
módulo no ejecuta ``git``: cada prueba le inyecta su runner.

stdlib exclusivamente, ``unittest``, repos ``git`` reales en ``tempfile``.
"""

from __future__ import annotations

import importlib.util
import os
import subprocess
import tempfile
import unittest
from pathlib import Path

BUDGET = Path(__file__).resolve().parents[1] / "vendored_budget.py"


def _load_budget():
    spec = importlib.util.spec_from_file_location("vendored_budget", BUDGET)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


budget = _load_budget()


def _git(repo: Path, *args: str) -> str:
    env = {k: v for k, v in os.environ.items() if not k.startswith("GIT_")}
    out = subprocess.run(
        ["git", "-c", "user.name=t", "-c", "user.email=t@example.invalid", *args],
        cwd=repo, capture_output=True, text=True, check=True, env=env,
    )
    return out.stdout


class TestVendoredRoots(unittest.TestCase):
    def test_the_only_root_is_the_vendored_tree(self) -> None:
        """#1088: solo queda fuera el árbol vendorizado de los consumers."""
        self.assertEqual(budget.VENDORED_ROOTS, (".agents/skills/",))

    def test_a_path_under_a_root_is_vendored(self) -> None:
        self.assertTrue(budget.is_under_vendored(".agents/skills/foo/x.py"))

    def test_the_catalog_is_not_vendored(self) -> None:
        """#1088: ``skills/`` es el catálogo de autoría del canon."""
        self.assertFalse(budget.is_under_vendored("skills/legacy/x.py"))
        self.assertFalse(budget.is_under_vendored("skills/acme/SKILL.md"))

    def test_an_authored_path_is_not_vendored(self) -> None:
        self.assertFalse(budget.is_under_vendored("app/auth.py"))
        self.assertFalse(budget.is_under_vendored("skills-notes/x.py"))


class TestCountsToBudget(unittest.TestCase):
    def test_authored_path_counts(self) -> None:
        self.assertTrue(budget.counts_to_budget("app/auth.py", "app/auth.py"))

    def test_vendored_path_without_rename_does_not_count(self) -> None:
        self.assertFalse(budget.counts_to_budget(".agents/skills/foo/x.py", ".agents/skills/foo/x.py"))

    def test_catalog_path_without_rename_counts(self) -> None:
        """#1088: bajo ``skills/`` es autoría y cuenta para el presupuesto."""
        self.assertTrue(budget.counts_to_budget("skills/legacy/x.py", "skills/legacy/x.py"))

    def test_rename_from_authored_into_vendored_counts(self) -> None:
        self.assertTrue(budget.counts_to_budget("app/auth.py", ".agents/skills/foo/auth.py"))
        self.assertTrue(budget.counts_to_budget("app/auth.py", "skills/legacy/auth.py"))

    def test_rename_within_vendored_does_not_count(self) -> None:
        self.assertFalse(budget.counts_to_budget(".agents/skills/a.py", ".agents/skills/b.py"))


class TestParseNumstat(unittest.TestCase):
    def test_plain_path(self) -> None:
        self.assertEqual(budget.parse_numstat("app/a.py"), ("app/a.py", "app/a.py"))

    def test_brace_rename_with_common_prefix_and_suffix(self) -> None:
        self.assertEqual(
            budget.parse_numstat("src/{old => .agents/skills/new}/x.py"),
            ("src/old/x.py", "src/.agents/skills/new/x.py"),
        )

    def test_brace_rename_into_a_vendored_root(self) -> None:
        self.assertEqual(
            budget.parse_numstat("{src => .agents/skills}/x.py"),
            ("src/x.py", ".agents/skills/x.py"),
        )

    def test_full_brace_rename(self) -> None:
        self.assertEqual(budget.parse_numstat("{app => skills}"), ("app", "skills"))

    def test_binary_or_empty_field_is_ignored(self) -> None:
        self.assertIsNone(budget.parse_numstat("-"))
        self.assertIsNone(budget.parse_numstat(""))


class TestBudgetLines(unittest.TestCase):
    NUMSTAT = (
        "10\t2\tapp/a.py\n"
        "300\t0\t.agents/skills/prop.md\n"
        "5\t5\tsrc/{old => .agents/skills/new}/x.py\n"
        "-\t-\tlogo.png\n"
    )

    def test_without_origins_everything_counts(self) -> None:
        self.assertEqual(budget.budget_lines(self.NUMSTAT), 10 + 2 + 300 + 5 + 5)

    def test_vendored_file_with_authored_origin_counts(self) -> None:
        origins = {".agents/skills/prop.md": "app/prop.md"}
        self.assertEqual(budget.budget_lines(self.NUMSTAT, origins), 10 + 2 + 300 + 5 + 5)

    def test_vendored_file_with_vendored_origin_is_excluded(self) -> None:
        origins = {".agents/skills/prop.md": ".agents/skills/prop.md"}
        self.assertEqual(budget.budget_lines(self.NUMSTAT, origins), 10 + 2 + 5 + 5)

    def test_vendored_file_missing_from_origins_counts(self) -> None:
        self.assertEqual(budget.budget_lines(self.NUMSTAT, {}), 10 + 2 + 300 + 5 + 5)

    def test_rename_from_authored_stays_counted_with_origins(self) -> None:
        origins = {".agents/skills/new/x.py": None}
        self.assertEqual(budget.budget_lines(self.NUMSTAT, origins), 10 + 2 + 5 + 5 + 300)


class TestVendoredLines(unittest.TestCase):
    """#1133: la cifra del árbol vendorizado, para que el gate diga qué parte
    de la cuenta no es trabajo de quien abre la PR."""

    NUMSTAT = (
        "10\t2\tapp/a.py\n"
        "300\t0\t.agents/skills/prop.md\n"
        "5\t5\tsrc/{old => .agents/skills/new}/x.py\n"
        "-\t-\tlogo.png\n"
    )

    def test_it_sums_the_lines_whose_destination_is_vendored(self) -> None:
        # El renombrado del fixture apunta a `src/.agents/skills/new/x.py`: no cae
        # bajo la raíz vendorizada, así que no cuenta como vendorizado.
        self.assertEqual(budget.vendored_lines(self.NUMSTAT), 300)

    def test_a_numstat_without_vendored_paths_is_zero(self) -> None:
        self.assertEqual(budget.vendored_lines("10\t2\tapp/a.py\n"), 0)

    def test_the_authored_catalog_is_not_vendored(self) -> None:
        """#1088: `skills/` es el catálogo de autoría del canon."""
        self.assertEqual(budget.vendored_lines("7\t0\tskills/acme/SKILL.md\n"), 0)


class TestParseOrigins(unittest.TestCase):
    def test_dash_means_unknown_origin(self) -> None:
        self.assertEqual(
            budget.parse_origins(".agents/skills/a.md\t-\n.agents/skills/b.md\tapp/b.md\n"),
            {".agents/skills/a.md": None, ".agents/skills/b.md": "app/b.md"},
        )

    def test_empty_text_gives_no_origins(self) -> None:
        self.assertEqual(budget.parse_origins(""), {})


class TestOriginsFor(unittest.TestCase):
    def _run_git(self, repo: Path) -> object:
        """Runner ligado al repo, con el contrato del módulo: ``RuntimeError`` si git falla."""
        return lambda *args: _git(repo, *args)

    def test_cross_commit_rename_reports_the_authored_origin(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            _git(repo, "init", "-q", "-b", "main")
            (repo / "app").mkdir()
            (repo / "app/a.py").write_text("".join(f"l{i}\n" for i in range(50)), encoding="utf-8")
            _git(repo, "add", "-A")
            _git(repo, "commit", "-q", "-m", "chore: base")
            _git(repo, "checkout", "-q", "-b", "feat/x")
            (repo / "src").mkdir()
            (repo / "src/x.py").write_text("".join(f"l{i}\n" for i in range(300)), encoding="utf-8")
            _git(repo, "add", "-A")
            _git(repo, "commit", "-q", "-m", "feat: autoria")
            (repo / ".agents/skills").mkdir(parents=True)
            _git(repo, "mv", "src/x.py", ".agents/skills/x.py")
            _git(repo, "commit", "-q", "-m", "chore(fleet): propagar")
            self.assertEqual(budget.origins_for("main", self._run_git(repo)),
                             ".agents/skills/x.py\tsrc/x.py\n")

    def test_no_vendored_files_gives_empty_text(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            _git(repo, "init", "-q", "-b", "main")
            (repo / "a.py").write_text("x\n", encoding="utf-8")
            _git(repo, "add", "-A")
            _git(repo, "commit", "-q", "-m", "chore: base")
            _git(repo, "checkout", "-q", "-b", "feat/x")
            (repo / "b.py").write_text("y\n", encoding="utf-8")
            _git(repo, "add", "-A")
            _git(repo, "commit", "-q", "-m", "feat: b")
            self.assertEqual(budget.origins_for("main", self._run_git(repo)), "")

    def test_a_failing_runner_leaves_the_origin_unknown(self) -> None:
        """Pin: un `git log --follow` que falla no rompe el presupuesto; cuenta hacia dentro."""
        def failing(*args: str) -> str:
            raise RuntimeError("git log failed")

        self.assertIsNone(budget.original_path(".agents/skills/x.py", "main", failing))

    def test_a_failing_runner_renders_a_dash_origin(self) -> None:
        def failing(*args: str) -> str:
            if args[0] == "log":
                raise RuntimeError("git log failed")
            return ".agents/skills/x.py\n"

        self.assertEqual(budget.origins_for("main", failing), ".agents/skills/x.py\t-\n")


if __name__ == "__main__":
    unittest.main()
