#!/usr/bin/env python3
"""Suite del asset de métricas de fricción de gates (#1179, fase 0).

Las cifras que decidirán el flip de la fase 2 (rojos secundarios, gates rojos
por commit, re-ejecuciones y tiempo hasta el verde) se prueban con fixtures que
violan una regla, más el contrato de salida (HR-3: una medición que no midió
sale 2, nunca 0). Sin red: el asset lee un fichero.
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "measure_gates.py"


def check(name, conclusion, started="2026-10-24T10:00:00Z",
          completed="2026-10-24T10:05:00Z"):
    """Un check-run del export."""
    return {"name": name, "conclusion": conclusion,
            "started_at": started, "completed_at": completed}


def commit(checks, pr=1173, sha="abc1234"):
    """Un commit del export."""
    return {"sha": sha, "pr": pr, "committed_at": "2026-10-24T09:59:00Z", "checks": checks}


def export(commits, repo="DysTelefonica/team-skills"):
    """El export completo."""
    return {"repo": repo, "window": {"from": "2026-10-17T00:00:00Z",
                                     "to": "2026-10-24T00:00:00Z"},
            "commits": commits}


def run_gate(doc, raw=None):
    """Corre el asset sobre un export; devuelve (rc, stdout+stderr)."""
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "export.json"
        path.write_text(raw if raw is not None else json.dumps(doc), encoding="utf-8")
        proc = subprocess.run(
            [sys.executable, str(GATE), "--export", str(path)],
            capture_output=True, text=True, check=False)
    return proc.returncode, proc.stdout + proc.stderr


class MetricRulesTests(unittest.TestCase):
    """Las reglas de cada cifra, con un fixture que viola una."""

    def assert_out(self, doc, needle):
        """El asset sale 0 y su salida nombra la regla comprobada."""
        code, out = run_gate(doc)
        self.assertEqual(code, 0, out)
        self.assertIn(needle, out)

    def test_one_red_is_not_secondary(self):
        self.assert_out(export([commit([check("unit", "failure")])]),
                        "rojos secundarios: 0")

    def test_two_reds_are_one_secondary_red(self):
        self.assert_out(export([commit([
            check("unit", "failure"), check("pr-size", "failure")])]),
            "rojos secundarios: 1")

    def test_the_aggregate_does_not_count_as_red(self):
        # El agregador `required` repite el rojo de la hoja: contarlo sería
        # medir dos veces el mismo defecto.
        self.assert_out(export([commit([
            check("unit", "failure"), check("required", "failure")])]),
            "rojos secundarios: 0")

    def test_cancelled_and_skipped_are_not_red(self):
        self.assert_out(export([commit([
            check("unit", "cancelled"), check("full suite", "skipped")])]),
            "commits con rojo: 0")

    def test_the_worst_commit_is_named(self):
        code, out = run_gate(export([commit([
            check("unit", "failure"), check("pr-size", "failure"),
            check("propagate", "failure")], pr=1103, sha="f5c2811f")]))
        self.assertEqual(code, 0, out)
        self.assertIn("peor commit: 3 checks rojos", out)
        self.assertIn("PR 1103", out)

    def test_a_second_run_of_the_same_check_is_a_rerun(self):
        self.assert_out(export([commit([
            check("unit", "failure", started="2026-10-24T10:00:00Z"),
            check("unit", "success", started="2026-10-24T10:20:00Z")])]),
            "commits con re-ejecución: 1")

    def test_a_single_run_is_not_a_rerun(self):
        self.assert_out(export([commit([check("unit", "success")])]),
                        "commits con re-ejecución: 0")

    def test_the_delay_is_measured_from_the_first_run_to_the_last_green(self):
        self.assert_out(export([commit([
            check("unit", "success", started="2026-10-24T10:00:00Z",
                  completed="2026-10-24T10:04:00Z"),
            check("pr-size", "success", started="2026-10-24T10:00:10Z",
                  completed="2026-10-24T10:06:00Z")])]), "mediana 6.0")

    def test_a_later_wave_is_not_the_wait_of_the_merge(self):
        # Un run que empieza más tarde que la ventana de la onda es otra
        # espera (re-ejecución o job filtrado por paths).
        self.assert_out(export([commit([
            check("unit", "success", started="2026-10-24T10:00:00Z",
                  completed="2026-10-24T10:04:00Z"),
            check("propagate", "success", started="2026-10-24T23:00:00Z",
                  completed="2026-10-24T23:09:00Z")])]), "mediana 4.0")

    def test_the_json_output_carries_the_metrics(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "export.json"
            path.write_text(json.dumps(export([commit([
                check("unit", "failure"), check("pr-size", "failure")])])),
                encoding="utf-8")
            proc = subprocess.run(
                [sys.executable, str(GATE), "--export", str(path), "--json"],
                capture_output=True, text=True, check=False)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        metrics = json.loads(proc.stdout)
        self.assertEqual(metrics["secondary_reds"], 1)
        self.assertEqual(metrics["red_commits"], 1)
        self.assertEqual(metrics["worst_commit"]["reds"], 2)


class InvalidExportTests(unittest.TestCase):
    """HR-3: un export ilegible o vacío sale 2, nunca 0."""

    def test_every_invalid_export_exits_two(self):
        cases = [("sin commits", json.dumps(export([])), "no tiene commits"),
                 ("sin la clave commits", json.dumps({"repo": "x"}), "'commits'"),
                 ("json roto", "{ esto no es json", "ilegible")]
        for label, raw, needle in cases:
            with self.subTest(case=label):
                code, out = run_gate(None, raw=raw)
                self.assertEqual(code, 2, out)
                self.assertIn(needle, out)

    def test_a_missing_file_exits_two(self):
        proc = subprocess.run(
            [sys.executable, str(GATE), "--export", "/nonexistent/export.json"],
            capture_output=True, text=True, check=False)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
