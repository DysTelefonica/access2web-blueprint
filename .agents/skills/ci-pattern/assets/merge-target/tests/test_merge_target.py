#!/usr/bin/env python3
# ci-pattern asset tests — merge-target gate (DysTelefonica/team-skills#420)
"""Suite del gate de base de merge (HR-59): la base de una PR es la rama por
defecto, salvo cadena valida (`chain:partial` y base = head de otra PR
abierta). Black-box: payload JSON y lista de PRs abiertas como ficheros; sin
API y sin git. Ejecucion: `python3 test_merge_target.py`."""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_merge_target.py"


class Base(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory(prefix="merge-target.")
        self.addCleanup(self._tmp.cleanup)
        self.root = Path(self._tmp.name)

    def event(self, base_ref, labels=("chain:partial",), number=9):
        p = self.root / f"event-{base_ref.replace('/', '-')}-{number}.json"
        p.write_text(json.dumps({"pull_request": {
            "number": number,
            "base": {"ref": base_ref},
            "labels": [{"name": item} for item in labels],
        }}), encoding="utf-8")
        return str(p)

    def open_heads(self, *refs):
        p = self.root / "open-heads.txt"
        p.write_text("\n".join(refs) + "\n", encoding="utf-8")
        return str(p)

    def run_gate(self, event, heads, *extra):
        return subprocess.run(
            [sys.executable, str(GATE), "--event", event,
             "--open-prs", heads, *extra],
            capture_output=True, text=True)


class HappyPathTests(Base):
    def test_base_default_is_clean(self):
        proc = self.run_gate(self.event("main", ()), self.open_heads("feat/x"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("rama por defecto", proc.stdout)

    def test_valid_chain_is_clean(self):
        heads = self.open_heads("feat/tramo-anterior", "feat/otra")
        proc = self.run_gate(self.event("feat/tramo-anterior"), heads)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("cadena valida", proc.stdout)

    def test_json_variant_of_open_prs_accepted(self):
        p = self.root / "open-heads.json"
        p.write_text(json.dumps([{"headRefName": "feat/tramo-anterior"}]),
                     encoding="utf-8")
        proc = self.run_gate(self.event("feat/tramo-anterior"), str(p))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


class FindingTests(Base):
    def test_wrong_base_without_chain_label_is_red(self):
        # El hallazgo de auditoria original: PR apilada mergeada en su rama
        # base sin declarar cadena.
        proc = self.run_gate(self.event("feat/huesped", ()),
                             self.open_heads("feat/huesped", "feat/otro"))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("sin la etiqueta de cadena", proc.stderr)
        self.assertIn("reabrir contra 'main'", proc.stderr)

    def test_chain_label_with_unopened_base_is_red(self):
        proc = self.run_gate(self.event("feat/cerrada"),
                             self.open_heads("feat/abierta"))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("rama huesped o ya cerrada", proc.stderr)

    def test_default_branch_is_a_parameter(self):
        # Consumidor con rama activa distinta: main pasa a ser base rara.
        proc = self.run_gate(self.event("main", ()),
                             self.open_heads("feat/x"),
                             "--default-branch", "release/2026")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_chain_label_configurable(self):
        proc = self.run_gate(self.event("feat/base", ("chain-apilada",)),
                             self.open_heads("feat/base"),
                             "--chain-label", "chain-apilada")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


class FailClosedTests(Base):
    def test_missing_payload_is_invalid(self):
        proc = self.run_gate(str(self.root / "nope.json"),
                             self.open_heads("feat/x"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_malformed_payload_is_invalid(self):
        p = self.root / "bad.json"
        p.write_text("{esto no es json", encoding="utf-8")
        proc = self.run_gate(str(p), self.open_heads("feat/x"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_payload_without_pull_request_is_invalid(self):
        p = self.root / "sin-pr.json"
        p.write_text(json.dumps({"issue": {}}), encoding="utf-8")
        proc = self.run_gate(str(p), self.open_heads("feat/x"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_payload_without_base_is_invalid(self):
        p = self.root / "sin-base.json"
        p.write_text(json.dumps({"pull_request": {"number": 1, "labels": []}}),
                     encoding="utf-8")
        proc = self.run_gate(str(p), self.open_heads("feat/x"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_empty_open_prs_list_is_invalid(self):
        # Lista vacia = fallo de la API previa, no cero PRs abiertas (HR-3).
        p = self.root / "heads-vacio.txt"
        p.write_text("", encoding="utf-8")
        proc = self.run_gate(self.event("main", ()), str(p))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_missing_open_prs_file_is_invalid(self):
        proc = self.run_gate(self.event("main", ()),
                             str(self.root / "nope.txt"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
