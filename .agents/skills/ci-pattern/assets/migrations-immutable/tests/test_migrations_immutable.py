#!/usr/bin/env python3
# ci-pattern asset tests — migrations-immutable, parte 1: estructura (#431)
"""Tests de estructura del gate migrations-immutable.

Cubre los criterios del issue #431 sobre la integridad de las
migraciones presentes en la rama base: rename, delete y edit son
rojos. Los tests de migraciones nuevas (prefijo mayor, sin prefijo,
duplicado nuevo, duplicado historico tolerado) viven en
test_migrations_immutable_new.py para mantener este archivo bajo
el presupuesto de revision.

Helpers compartidos en _helpers.py.
"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _helpers import (
    BaseMigrationTest, add_files, git, make_repo, run_gate,
)


class StructureTest(BaseMigrationTest):
    """#431 parte 1: rename, delete, edit sobre migraciones existentes."""

    def test_no_migrations_is_clean(self) -> None:
        repo = make_repo(self.tmp, {"README.md": "hi"})
        proc = run_gate(repo, "--base", "HEAD")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_rename_existing_migration_is_red(self) -> None:
        repo = make_repo(self.tmp, {
            "migrations/001_init.sql": "CREATE TABLE a;",
        })
        (repo / "migrations/001_init.sql").rename(
            repo / "migrations/001_renamed.sql"
        )
        git(repo, "add", "-A")
        git(repo, "commit", "-q", "-m", "rename")
        proc = run_gate(repo, "--base", "HEAD~1")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("rename", proc.stdout.lower())

    def test_delete_existing_migration_is_red(self) -> None:
        repo = make_repo(self.tmp, {
            "migrations/001_init.sql": "CREATE TABLE a;",
        })
        (repo / "migrations/001_init.sql").unlink()
        git(repo, "add", "-A")
        git(repo, "commit", "-q", "-m", "delete")
        proc = run_gate(repo, "--base", "HEAD~1")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("delete", proc.stdout.lower())

    def test_edit_existing_migration_is_red(self) -> None:
        repo = make_repo(self.tmp, {
            "migrations/001_init.sql": "CREATE TABLE a;",
        })
        (repo / "migrations/001_init.sql").write_text(
            "CREATE TABLE a; ALTER TABLE a ADD COLUMN x INT;",
            encoding="utf-8",
        )
        git(repo, "add", "-A")
        git(repo, "commit", "-q", "-m", "edit")
        proc = run_gate(repo, "--base", "HEAD~1")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("edit", proc.stdout.lower())

    def test_no_git_repo_is_fail_closed(self) -> None:
        not_repo = self.tmp / "not_repo"
        not_repo.mkdir()
        proc = run_gate(not_repo, "--base", "HEAD")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_unresolvable_base_is_fail_closed(self) -> None:
        repo = make_repo(self.tmp, {"README.md": "hi"})
        proc = run_gate(repo, "--base", "this-ref-does-not-exist")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main()
