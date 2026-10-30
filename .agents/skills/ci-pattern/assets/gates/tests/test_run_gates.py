#!/usr/bin/env python3
"""Suite del job de gates con veredicto único, en sombra (#1179, fase 1).

El job ejecuta los gates declarados y agrega sus veredictos en uno. Aquí se
prueban las reglas: todos en verde, alguno en rojo (con su causa citada), un
gate que no midió (timeout o crash) y una declaración rota (HR-3: un gate que
no midió nunca dice «OK»).

Ejecuta sin red: los gates del fixture son comandos locales triviales.
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "run_gates.py"


def write_gates(tmp: str, body: str) -> str:
    """El fichero de declaración de gates."""
    path = Path(tmp) / "gates.txt"
    path.write_text(body, encoding="utf-8")
    return str(path)


def run_job(tmp: str, body: str, *extra) -> tuple:
    """Corre el job con la declaración dada; devuelve (rc, stdout+stderr)."""
    gates = write_gates(tmp, body)
    proc = subprocess.run(
        [sys.executable, str(GATE), "--gates", gates, "--cwd", tmp, *extra],
        capture_output=True, text=True, check=False)
    return proc.returncode, proc.stdout + proc.stderr


class VerdictTests(unittest.TestCase):
    """El veredicto único: cuántos en verde y por qué fallan los que fallan."""

    def test_all_gates_green_exits_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, out = run_job(tmp, "a: exit 0\nb: exit 0\n")
        self.assertEqual(code, 0, out)
        self.assertIn("gates: 2/2 en verde", out)

    def test_a_failing_gate_exits_one_and_names_its_cause(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, out = run_job(tmp, "ok: exit 0\nko: echo 'el gate fallo por X'; exit 3\n")
        self.assertEqual(code, 1, out)
        self.assertIn("gates: 1/2 en verde", out)
        self.assertIn("ko", out)
        self.assertIn("el gate fallo por X", out)
        self.assertIn("exit 3", out)

    def test_a_gate_that_did_not_measure_exits_two(self):
        # Un gate que no termina no es un rojo: es una ausencia (HR-3).
        with tempfile.TemporaryDirectory() as tmp:
            code, out = run_job(tmp, "lento: sleep 5\n", "--timeout", "1")
        self.assertEqual(code, 2, out)
        self.assertIn("no midió", out)
        self.assertIn("lento", out)

    def test_the_json_output_carries_the_verdict(self):
        with tempfile.TemporaryDirectory() as tmp:
            gates = write_gates(tmp, "ok: exit 0\nko: exit 1\n")
            proc = subprocess.run(
                [sys.executable, str(GATE), "--gates", gates, "--cwd", tmp, "--json"],
                capture_output=True, text=True, check=False)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        doc = json.loads(proc.stdout)
        self.assertEqual(doc["gates"], 2)
        self.assertEqual(doc["passed"], 1)
        self.assertEqual(doc["failed"], 1)
        self.assertEqual(doc["failing"][0]["name"], "ko")


class DeclarationTests(unittest.TestCase):
    """Una declaración rota sale 2: sin gates declarados el job no mide nada."""

    def test_an_empty_declaration_exits_two(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, out = run_job(tmp, "# solo un comentario\n")
        self.assertEqual(code, 2, out)
        self.assertIn("vacía", out)

    def test_a_line_without_a_name_exits_two(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, out = run_job(tmp, "sin dos puntos\n")
        self.assertEqual(code, 2, out)
        self.assertIn("necesita un nombre", out)

    def test_an_invalid_name_exits_two(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, out = run_job(tmp, "nombre con espacios: exit 0\n")
        self.assertEqual(code, 2, out)
        self.assertIn("no es un nombre de gate válido", out)

    def test_a_gate_without_a_command_exits_two(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, out = run_job(tmp, "sin-comando:\n")
        self.assertEqual(code, 2, out)
        self.assertIn("no declara comando", out)

    def test_a_missing_file_exits_two(self):
        with tempfile.TemporaryDirectory() as tmp:
            proc = subprocess.run(
                [sys.executable, str(GATE), "--gates", f"{tmp}/no-existe.txt"],
                capture_output=True, text=True, check=False)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("declaración de gates rota", proc.stdout + proc.stderr)


class ParsingTests(unittest.TestCase):
    """La declaración se lee como `nombre: comando`, y el resto es del comando."""

    def test_comments_and_blank_lines_are_ignored(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, out = run_job(tmp, "# comentario\n\ngate: exit 0\n")
        self.assertEqual(code, 0, out)
        self.assertIn("gates: 1/1 en verde", out)

    def test_the_command_keeps_its_colons(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, out = run_job(tmp, "echo: echo 'a:b:c'\n")
        self.assertEqual(code, 0, out)


if __name__ == "__main__":
    unittest.main(verbosity=2)
