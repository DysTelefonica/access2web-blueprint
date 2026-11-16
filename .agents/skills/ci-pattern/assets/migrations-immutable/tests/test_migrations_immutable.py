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
import subprocess
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _helpers import (
    BaseMigrationTest,
    git,
    make_repo,
    run_gate,
    run_gate_under_cp1252_locale,
)


class StructureTest(BaseMigrationTest):
    """#431 parte 1: rename, delete, edit sobre migraciones existentes."""

    def test_invalid_utf8_rename_fails_closed(self) -> None:
        """#1601: two non-UTF-8 names must not collapse into one key and pass."""
        repo = make_repo(self.tmp, {"migrations/001_init.sql": "CREATE TABLE a;"})
        git(repo, "config", "core.quotepath", "false")

        def run(*a, inp=None):
            return subprocess.run(["git", "-C", str(repo), *a], input=inp,
                                  capture_output=True, check=True).stdout.decode().strip()

        blob = run("hash-object", "-w", "--stdin", inp=b"CREATE TABLE b;")
        first = run("rev-parse", "HEAD:migrations/001_init.sql")

        def commit(name, parent):
            sub = run("mktree", inp=b"100644 blob " + first.encode() + b"\t001_init.sql\n"
                      + b"100644 blob " + blob.encode() + b"\t" + name + b"\n")
            top = run("mktree", inp=b"040000 tree " + sub.encode() + b"\tmigrations\n")
            c = run("commit-tree", top, "-p", parent, "-m", "x")
            run("update-ref", "HEAD", c)
            return c

        base = commit(b"002_\xf3.sql", run("rev-parse", "HEAD"))
        commit(b"002_\xe1.sql", base)
        code, out, err = run_gate_under_cp1252_locale(repo, "--base", base)
        self.assertNotEqual(code, 0, out + err)
        self.assertIn("no es UTF-8", out + err)

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

    def test_non_ascii_migration_path_is_read_under_cp1252(self) -> None:
        """#1601: the migration paths from git are UTF-8 whatever the locale."""
        repo = make_repo(self.tmp, {"migrations/001_init.sql": "CREATE TABLE a;"})
        git(repo, "config", "core.quotepath", "false")
        (repo / "migrations/002_sección_Á.sql").write_text("CREATE TABLE b;", encoding="utf-8")
        git(repo, "add", "-A")
        git(repo, "commit", "-q", "-m", "add")
        code, out, err = run_gate_under_cp1252_locale(repo, "--base", "HEAD~1")
        self.assertEqual(code, 0, out + err)


if __name__ == "__main__":
    unittest.main()
