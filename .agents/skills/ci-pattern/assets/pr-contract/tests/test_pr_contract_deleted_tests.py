#!/usr/bin/env python3
"""HR-88 (#1509): la PR que borra `def test_` de un módulo que sigue en uso avisa
(`::warning::`, exit 0) salvo que el cuerpo diga dónde está ahora la cobertura.
stdlib exclusivamente."""

import os
import subprocess
import tempfile
import unittest
from pathlib import Path

import test_pr_contract as gate

SERVICE = (
    "def list_items(q):\n    return []\n\n\n"
    "def search_items(q):\n    return [q.replace('%', '\\\\%')]\n")
ROUTES = (
    "from app.service import list_items, search_items\n\n\n"
    "def index():\n    return list_items(''), search_items('')\n")
TEST_SERVICE = (
    "from app.service import list_items, search_items\n\n\n"
    "def test_list_items():\n    assert list_items('') == []\n\n\n"
    "def test_search_escapes_like():\n    assert search_items('%') == ['\\\\%']\n")
TEST_SERVICE_WITHOUT_SEARCH = (
    "from app.service import list_items\n\n\n"
    "def test_list_items():\n    assert list_items('') == []\n")
BASE_TREE = {"app/__init__.py": "", "app/service.py": SERVICE, "app/routes.py": ROUTES,
             "tests/test_service.py": TEST_SERVICE}


def make_tree_repo(tmp, base_files, head_changes):
    """Repo con `base_files` en main y `head_changes` (None borra) en feature/x."""
    repo = Path(tmp) / "repo"
    repo.mkdir()
    env = {k: v for k, v in os.environ.items() if not k.startswith("GIT_")}

    def run(*args):
        return subprocess.run(args, cwd=repo, capture_output=True, text=True, check=True,
                              env={**env, "GIT_AUTHOR_NAME": "t", "GIT_AUTHOR_EMAIL": "t@t",
                                   "GIT_COMMITTER_NAME": "t", "GIT_COMMITTER_EMAIL": "t@t"})

    def commit(files, message):
        for rel, text in files.items():
            path = repo / rel
            if text is None:
                path.unlink()
            else:
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(text, encoding="utf-8", newline="\n")
        run("git", "add", "-A")
        run("git", "commit", "-qm", message)
        return run("git", "rev-parse", "HEAD").stdout.strip()

    run("git", "init", "-q", "-b", "main")
    base_sha = commit(base_files, "chore: base")
    run("git", "checkout", "-q", "-b", "feature/x")
    return repo, base_sha, commit(head_changes, "test: reorganiza los tests")


def run_deleted_tests(head_changes, *, body="", base_files=None, with_diff=True, checks=None):
    """Ejecuta el gate sobre la PR `head_changes`, con el diff como lo genera el workflow."""
    with tempfile.TemporaryDirectory() as tmp:
        repo, base_sha, head_sha = make_tree_repo(
            tmp, BASE_TREE if base_files is None else base_files, head_changes)
        ev = gate.make_event(tmp, body=body + "\n" + gate.AI_OK, labels=["type:feature"],
                             base_sha=base_sha, head_sha=head_sha)
        pol = gate.make_policy(tmp, **({} if checks is None else {"checks": checks}))
        with tempfile.TemporaryDirectory() as dest:
            flags = gate.write_range_data(ev, repo, dest)
            if with_diff:
                diff = Path(dest) / "diff.txt"
                diff.write_text(gate.git_text(repo, "diff", "--no-renames", "-U0",
                                              f"{base_sha}...{head_sha}", "--", "*.py"),
                                encoding="utf-8")
                flags += ["--diff-file", str(diff)]
            return gate.run_gate(ev, pol, repo, data_flags=flags)


def warnings_of(proc):
    return [line for line in proc.stdout.splitlines() if line.startswith("::warning::")]


class DeletedTestsTests(unittest.TestCase):
    """HR-88 (#1509): tests borrados de un módulo que sigue en uso."""

    def test_deleted_test_of_a_module_in_use_warns_naming_test_and_module(self):
        p = run_deleted_tests({"tests/test_service.py": TEST_SERVICE_WITHOUT_SEARCH})
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        warnings = warnings_of(p)
        self.assertEqual(len(warnings), 1, p.stdout)
        for part in ("HR-88", "tests/test_service.py::test_search_escapes_like",
                     "app.service", "app/routes.py", "Cobertura movida"):
            self.assertIn(part, warnings[0])
        self.assertNotIn("test_list_items", warnings[0])

    def test_empty_moved_coverage_section_still_warns(self):
        body = "## Cobertura movida\n\n<!-- dónde está ahora la cobertura -->\n\n## Otra\n\nx\n"
        p = run_deleted_tests({"tests/test_service.py": TEST_SERVICE_WITHOUT_SEARCH}, body=body)
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertEqual(len(warnings_of(p)), 1, p.stdout)

    def test_non_empty_moved_coverage_section_silences_the_warning(self):
        body = ("## Cobertura movida\n\n"
                "- `tests/test_api.py::test_search` cubre el escape de LIKE.\n")
        p = run_deleted_tests({"tests/test_service.py": TEST_SERVICE_WITHOUT_SEARCH}, body=body)
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertEqual(warnings_of(p), [], p.stdout)

    def test_test_deleted_with_its_module_does_not_warn(self):
        p = run_deleted_tests({"tests/test_service.py": None, "app/service.py": None,
                               "app/routes.py": "def index():\n    return []\n"})
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertEqual(warnings_of(p), [], p.stdout)

    def test_test_deleted_with_the_function_it_exercised_does_not_warn(self):
        p = run_deleted_tests({
            "tests/test_service.py": TEST_SERVICE_WITHOUT_SEARCH,
            "app/service.py": "def list_items(q):\n    return []\n",
            "app/routes.py": ("from app.service import list_items\n\n\n"
                              "def index():\n    return list_items('')\n")})
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertEqual(warnings_of(p), [], p.stdout)

    def test_it_does_not_count_as_the_only_active_check(self):
        """HR-3: sin `--diff-file` no mide nada; con el resto apagado el gate sale 2."""
        checks = {name: name == "deleted_tests" for name in gate.ALL_CHECKS}
        p = run_deleted_tests({"tests/test_service.py": TEST_SERVICE_WITHOUT_SEARCH},
                              checks=checks)
        self.assertEqual(p.returncode, 2, p.stdout + p.stderr)
        self.assertIn("ninguna", p.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
