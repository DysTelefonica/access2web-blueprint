#!/usr/bin/env python3
"""Suite del generador del fichero de origenes del PR (#931, #1133).

El mapa de origenes dice que fichero vendorizado del PR llega desde donde: un
fichero nuevo bajo `.agents/skills/` no tiene origen de autoria (no cuenta para
el presupuesto) y un renombrado desde una ruta de autoria si (cuenta). El gate
`pr-contract` lo lee con `--origins-file`; sin el, cuenta todo.

El asset no ejecuta git (ratchet extendido): el workflow le pasa la salida de
git, asi que aqui se prueban las entradas ya capturadas.
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSETS = Path(__file__).resolve().parents[2]
SCRIPT = ASSETS / "bin" / "origins_file.py"
TEMPLATE = ASSETS / "templates" / ".github" / "workflows" / "pr-contract.yml.tmpl"


def _run(names: str, log: str, tmp: str) -> str:
    names_file = Path(tmp) / "names"
    log_file = Path(tmp) / "log"
    out = Path(tmp) / "origins"
    names_file.write_text(names, encoding="utf-8")
    log_file.write_text(log, encoding="utf-8")
    done = subprocess.run(
        [sys.executable, str(SCRIPT), "--base", "main", "--head", "HEAD",
         "--names-file", str(names_file), "--log-file", str(log_file),
         "--output", str(out)],
        capture_output=True, text=True, check=False)
    assert done.returncode == 0, done.stdout + done.stderr
    return out.read_text(encoding="utf-8")


class OriginsFileTest(unittest.TestCase):
    """El script escribe el mapa que el gate lee."""

    def test_a_new_vendored_file_has_no_authored_origin(self) -> None:
        """Un fichero nuevo bajo `.agents/skills/` no tiene origen: no cuenta."""
        with tempfile.TemporaryDirectory() as tmp:
            names = ".agents/skills/demo/SKILL.md\n"
            log = ""
            self.assertEqual(_run(names, log, tmp), ".agents/skills/demo/SKILL.md\t-\n")

    def test_a_rename_from_an_authored_path_keeps_its_origin(self) -> None:
        """Un renombrado desde una ruta de autoria conserva el origen: cuenta."""
        with tempfile.TemporaryDirectory() as tmp:
            names = ".agents/skills/demo/main.py\n"
            log = "R100\tapp/main.py\t.agents/skills/demo/main.py\n"
            self.assertEqual(
                _run(names, log, tmp),
                ".agents/skills/demo/main.py\tapp/main.py\n")

    def test_a_file_outside_the_vendored_roots_is_not_in_the_map(self) -> None:
        """Lo que no es vendorizado no entra en el mapa."""
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(_run("app/main.py\n", "", tmp), "")

    def test_the_gate_reads_the_file_with_the_origins_flag(self) -> None:
        """El gate acepta `--origins-file` y lo usa para el presupuesto."""
        gate = ASSETS / "pr-contract" / "check_pr_contract.py"
        self.assertIn("--origins-file", gate.read_text(encoding="utf-8"))


class TemplateTest(unittest.TestCase):
    """La plantilla de workflow del patron genera el fichero y lo pasa."""

    def test_the_template_generates_the_origins_file(self) -> None:
        self.assertTrue(TEMPLATE.is_file(), f"falta la plantilla {TEMPLATE}")
        text = TEMPLATE.read_text(encoding="utf-8")
        self.assertIn("origins_file.py", text)
        self.assertIn("--output", text)

    def test_the_template_passes_the_origins_file_to_pr_contract(self) -> None:
        text = TEMPLATE.read_text(encoding="utf-8")
        self.assertIn("--origins-file", text)
        self.assertIn("check_pr_contract.py", text)

    def test_the_template_computes_the_origins_over_the_merge_base_range(self) -> None:
        """El rango es `BASE...HEAD` (merge-base), como el resto de datos (#1013)."""
        text = TEMPLATE.read_text(encoding="utf-8")
        self.assertIn('"${BASE_SHA}...${HEAD_SHA}"', text)

    def test_the_template_runs_git_and_hands_the_output_to_the_script(self) -> None:
        """El git lo ejecuta el workflow: el asset no ejecuta git (ratchet)."""
        text = TEMPLATE.read_text(encoding="utf-8")
        # La cabeza del PR se checkoutea aparte (pr-head): el git corre ahí.
        self.assertIn("git -C pr-head diff --name-only", text)
        self.assertIn("git -C pr-head log --name-status", text)
        self.assertIn("--names-file", text)
        self.assertIn("--log-file", text)
        script = SCRIPT.read_text(encoding="utf-8")
        self.assertNotIn("import subprocess", script)


if __name__ == "__main__":
    unittest.main()
