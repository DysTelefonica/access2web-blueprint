#!/usr/bin/env python3
# ci-pattern asset tests — migrations-immutable, parte 2: nuevas (#431)
"""Tests de migraciones nuevas del gate migrations-immutable.

Cubre los criterios del issue #431 sobre las migraciones que se anaden
en HEAD: prefijo mayor al maximo de la base, sin prefijo, duplicado
nuevo rechazado, duplicado historico tolerado, adicion limpia.
Los tests de estructura (rename, delete, edit) viven en
test_migrations_immutable.py. Helpers compartidos en _helpers.py.
"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _helpers import (
    BaseMigrationTest, add_files, make_repo, run_gate,
)


class NewMigrationsTest(BaseMigrationTest):
    """#431 parte 2: validaciones sobre migraciones nuevas."""

    def test_clean_add_is_green(self) -> None:
        repo = make_repo(self.tmp, {
            "migrations/001_init.sql": "CREATE TABLE a;",
            "migrations/002_more.sql": "CREATE TABLE b;",
        })
        add_files(repo, {
            "migrations/003_next.sql": "CREATE TABLE c;",
        })
        proc = run_gate(repo, "--base", "HEAD~1")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("OK", proc.stdout)

    def test_duplicate_prefix_in_base_is_tolerated(self) -> None:
        # Duplicado historico en la base: el gate no lo reporta.
        repo = make_repo(self.tmp, {
            "migrations/001_init.sql": "CREATE TABLE a;",
            "migrations/001_legacy.sql": "CREATE TABLE legacy;",
        })
        add_files(repo, {
            "migrations/002_next.sql": "CREATE TABLE c;",
        })
        proc = run_gate(repo, "--base", "HEAD~1")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_new_with_lower_prefix_is_red(self) -> None:
        repo = make_repo(self.tmp, {
            "migrations/001_init.sql": "CREATE TABLE a;",
            "migrations/002_more.sql": "CREATE TABLE b;",
        })
        add_files(repo, {
            "migrations/001_collision.sql": "CREATE TABLE collision;",
        })
        proc = run_gate(repo, "--base", "HEAD~1")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("prefijo", proc.stdout.lower())

    def test_new_with_no_prefix_is_red(self) -> None:
        repo = make_repo(self.tmp, {
            "migrations/001_init.sql": "CREATE TABLE a;",
        })
        add_files(repo, {
            "migrations/unversioned.sql": "CREATE TABLE z;",
        })
        proc = run_gate(repo, "--base", "HEAD~1")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("sin prefijo", proc.stdout.lower())

    def test_new_duplicate_prefix_is_red(self) -> None:
        # Prefijo 001 ya duplicado en la base (001_init + 001_legacy).
        # Una nueva 001_otra debe ser rechazada.
        repo = make_repo(self.tmp, {
            "migrations/001_init.sql": "CREATE TABLE a;",
            "migrations/001_legacy.sql": "CREATE TABLE legacy;",
        })
        add_files(repo, {
            "migrations/001_otra.sql": "CREATE TABLE otra;",
        })
        proc = run_gate(repo, "--base", "HEAD~1")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("duplicado", proc.stdout.lower())


if __name__ == "__main__":
    unittest.main()
