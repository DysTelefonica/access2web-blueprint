#!/usr/bin/env python3
"""Tests del gate de ensayo nocturno (#1627, slice 8).

Todo PR que toque el pipeline nocturno necesita un ensayo nocturno en seco del
SHA de su cabeza antes de fusionarse. El gate solo lee: no escribe nada.

Códigos de salida: 0 si el PR puede fusionarse; 1 si el ensayo falta, no es
del SHA de la cabeza, no es en seco o no salió en verde; 2 si el registro del
ensayo no se puede leer (fallo cerrado).
"""

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "night_gates.py"
sys.path.insert(0, str(ASSET))
import night_gates  # noqa: E402

HEAD = "a" * 40
OTHER = "b" * 40
NIGHT_FILE = ".github/workflows/nightly-run.yml"


def run_gate(changed, record=None, head=HEAD):
    """Corre el gate en un directorio temporal y comprueba que no escribe nada."""
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        changed_file = root / "changed.txt"
        changed_file.write_text("\n".join(changed) + "\n", encoding="utf-8")
        args = [sys.executable, str(GATE), "dry-night", "--changed", str(changed_file),
                "--head", head]
        if record is not None:
            record_file = root / "record.json"
            record_file.write_text(record if isinstance(record, str) else json.dumps(record),
                                   encoding="utf-8")
            args += ["--record", str(record_file)]
        before = sorted(p.name for p in root.iterdir())
        proc = subprocess.run(args, capture_output=True, text=True, cwd=root)
        after = sorted(p.name for p in root.iterdir())
    return proc, before, after


class PathMatchTests(unittest.TestCase):
    def test_double_star_crosses_directories(self):
        self.assertTrue(night_gates.path_matches("a/b/c.sql", "**/*.sql"))
        self.assertTrue(night_gates.path_matches("c.sql", "**/*.sql"))

    def test_single_star_stays_in_one_directory(self):
        self.assertTrue(night_gates.path_matches("db/x.sql", "db/*.sql"))
        self.assertFalse(night_gates.path_matches("db/sub/x.sql", "db/*.sql"))

    def test_night_workflow_is_in_the_default_paths(self):
        self.assertTrue(night_gates.touches_night_pipeline([NIGHT_FILE]))
        self.assertFalse(night_gates.touches_night_pipeline([".github/workflows/tests.yml"]))

    def test_vendored_night_scripts_are_in_the_default_paths(self):
        for path in (".agents/skills/ci-pattern/assets/prerelease/night_gates.py",
                     ".agents/skills/ci-pattern/assets/prerelease/nightly_pipeline.py",
                     ".agents/skills/ci-pattern/assets/prerelease/night-tool.sh",
                     "personal/ardelperal/ci-pattern/assets/prerelease/night_gates.py"):
            with self.subTest(path=path):
                self.assertTrue(night_gates.touches_night_pipeline([path]))
        self.assertFalse(night_gates.touches_night_pipeline(
            [".agents/skills/ci-pattern/assets/prerelease/park_issues.py"]))


class DryNightGateTests(unittest.TestCase):
    def test_ordinary_pr_passes_without_any_record(self):
        proc, _, _ = run_gate(["src/app.py", "docs/README.md"])
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("no toca el pipeline nocturno", proc.stdout)

    def test_night_change_without_a_record_is_not_verified(self):
        proc, _, _ = run_gate([NIGHT_FILE])
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(NIGHT_FILE, proc.stderr)
        self.assertIn(HEAD, proc.stderr)

    def test_dry_green_run_for_this_head_passes(self):
        record = {"mode": "dry", "head_sha": HEAD, "result": "success"}
        proc, _, _ = run_gate([NIGHT_FILE], record=record)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_dry_run_of_another_sha_does_not_count(self):
        record = {"mode": "dry", "head_sha": OTHER, "result": "success"}
        proc, _, _ = run_gate([NIGHT_FILE], record=record)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(NIGHT_FILE, proc.stderr)

    def test_a_real_run_is_not_a_dry_run(self):
        record = {"mode": "real", "head_sha": HEAD, "result": "success"}
        proc, _, _ = run_gate([NIGHT_FILE], record=record)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_a_red_dry_run_does_not_pass(self):
        record = {"mode": "dry", "head_sha": HEAD, "result": "failure"}
        proc, _, _ = run_gate([NIGHT_FILE], record=record)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_an_unreadable_record_fails_closed_with_exit_2(self):
        proc, _, _ = run_gate([NIGHT_FILE], record="{no es json")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_the_gate_writes_nothing(self):
        record = {"mode": "dry", "head_sha": HEAD, "result": "success"}
        proc, before, after = run_gate([NIGHT_FILE], record=record)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(before, after)


class NightPathCoverageTests(unittest.TestCase):
    """Ninguna ruta del pipeline nocturno sale con 0 sin ensayo (fallo abierto, #1648)."""

    NIGHT_CHANGES = (
        ".github/workflows/nightly.yaml",
        "night-pipeline.yaml",
        "deploy-nightly.yml",
        "./.github/workflows/nightly.yml",
    )

    def test_every_declared_night_path_needs_a_dry_run(self):
        for path in self.NIGHT_CHANGES:
            with self.subTest(path=path):
                proc, _, _ = run_gate([path])
                self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_a_paths_file_declares_the_night_pipeline(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            paths = root / "paths.json"
            paths.write_text(json.dumps(["infra/cron-*.yml"]), encoding="utf-8")
            changed = root / "changed.txt"
            changed.write_text("infra/cron-nightly.yml\n", encoding="utf-8")
            proc = subprocess.run([sys.executable, str(GATE), "dry-night", "--changed", str(changed),
                                   "--head", HEAD, "--paths-file", str(paths)],
                                  capture_output=True, text=True, cwd=root, check=False)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_an_unreadable_paths_file_fails_closed(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            changed = root / "changed.txt"
            changed.write_text("src/app.py\n", encoding="utf-8")
            proc = subprocess.run([sys.executable, str(GATE), "dry-night", "--changed", str(changed),
                                   "--head", HEAD, "--paths-file", str(root / "no.json")],
                                  capture_output=True, text=True, cwd=root, check=False)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)


def run_with(changed_text: str, *extra: str, paths_json: str | None = None):
    """Corre el gate con un fichero --changed de texto libre y opciones extra."""
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        changed = root / "changed.txt"
        changed.write_text(changed_text, encoding="utf-8")
        args = [sys.executable, str(GATE), "dry-night", "--changed", str(changed), "--head", HEAD, *extra]
        if paths_json is not None:
            paths = root / "paths.json"
            paths.write_text(paths_json, encoding="utf-8")
            args += ["--paths-file", str(paths)]
        return subprocess.run(args, capture_output=True, text=True, cwd=root, check=False)


class RegressionTests(unittest.TestCase):
    """Regresiones del verificador: variantes de ruta que no deben salir con 0 (#1648)."""

    VARIANTS = (
        ".github/workflows/Nightly.yml",
        ".github/workflows/NIGHTLY.yml",
        ".github/workflows/nightly.YML",
        ".github\\workflows\\nightly.yml",
        "/.github/workflows/nightly.yml",
        ".//.github/workflows/nightly.yml",
    )

    def test_case_and_separator_variants_need_a_dry_run(self):
        for path in self.VARIANTS:
            with self.subTest(path=path):
                proc = run_with(path + "\n")
                self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_a_rename_away_from_a_nightly_file_needs_a_dry_run(self):
        """Con --no-renames el cambio trae la ruta vieja y la nueva: la vieja decide."""
        proc = run_with(".github/workflows/nightly.yml\n.github/workflows/ci.yml\n")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_an_empty_changed_list_fails_closed(self):
        for text in ("", "\n\n", "# solo comentario\n"):
            with self.subTest(text=text):
                proc = run_with(text)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_an_empty_paths_list_fails_closed(self):
        proc = run_with("src/app.py\n", paths_json="[]")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_night_path_adds_a_glob(self):
        proc = run_with("infra/cron.yml\n", "--night-path", "infra/*.yml")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)


    def test_uppercase_declared_globs_match(self):
        """Un glob declarado en mayúsculas casa con la ruta en minúsculas (#1648)."""
        proc = run_with("infra/cron.yml\n", paths_json=json.dumps(["INFRA/*.YML"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_inner_double_slash_and_mixed_separators(self):
        for path in (".github//workflows/nightly.yml", ".github\\workflows/nightly.yml",
                     ".github/workflows//nightly.yml"):
            with self.subTest(path=path):
                proc = run_with(path + "\n")
                self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_git_quoted_path_does_not_pass(self):
        """git cita las rutas no ASCII; el gate sale con 2 y dice el comando correcto."""
        proc = run_with('".github/workflows/nightly-\\303\\261.yml"\n')
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("core.quotePath=false", proc.stderr)

    def test_bom_on_first_line_does_not_hide_a_night_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            changed = root / "changed.txt"
            changed.write_bytes("﻿.github/workflows/nightly.yml\n".encode())
            proc = subprocess.run([sys.executable, str(GATE), "dry-night", "--changed", str(changed),
                                   "--head", HEAD], capture_output=True, text=True, cwd=root, check=False)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
