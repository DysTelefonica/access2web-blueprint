#!/usr/bin/env python3
# ci-pattern asset — el smoke-run del meta-gate HR→gate corre en su entorno (#342)
"""Suite del entorno con el que el meta-gate HR→gate ejecuta los tests
declarados (#342).

Esa ejecución NO es un gate anidado: el test puede llamar a `adoption` en
proceso (la suite real de HR-45/46/48 lo hace, para generar el documento
operativo). La marca `CI_PATTERN_NESTED` de la CLI (#277: la guarda que
evitó el fork bomb del 2026-10-05) se retira del smoke-run; el
encadenamiento se cierra por identidad —el test que ya está corriendo no se
vuelve a smoke-ejecutar, `CI_PATTERN_SMOKE_TESTS` lleva sus rutas— y la
profundidad (`CI_PATTERN_SMOKE`) queda como alambre trampa. Así ningún
camino recursa, y los tests que declaran los fixtures siguen juzgándose.

Exit 0 si el entorno y los topes son los declarados; 1 con hallazgos.
"""

import ast
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[3]
GATE = SKILL_ROOT / "assets" / "hr-gate-matrix" / "check_hr_matrix.py"
CLI = SKILL_ROOT / "assets" / "bin" / "ci-pattern"
MARKERS = ("CI_PATTERN_NESTED", "CI_PATTERN_SMOKE", "CI_PATTERN_SMOKE_TESTS")

# El test declarado por el mini-skill de los fixtures: comprueba el entorno
# del smoke-run y, si hay sonda, deja constancia de que se ejecutó.
PROBE_TEST = """
import os, sys

probe = os.environ.get("SMOKE_PROBE")
depth = os.environ.get("CI_PATTERN_SMOKE", "")
if probe:
    with open(probe, "a", encoding="utf-8") as handle:
        handle.write(f"ran depth={depth}\\n")
marker = os.environ.get("CI_PATTERN_NESTED")
if marker is not None or not depth.isdigit() or int(depth) < 1:
    print(f"entorno del smoke-run inesperado: marker={marker!r} depth={depth!r}", file=sys.stderr)
    sys.exit(1)
"""


def make_mini_skill(tmp, body=PROBE_TEST):
    """Mini-skill con una regla `gate` cuyo test declarado es `body`."""
    root = Path(tmp) / "skill"
    (root / "references").mkdir(parents=True)
    (root / "assets" / "gate" / "tests").mkdir(parents=True)
    (root / "SKILL.md").write_text("---\nname: mini\n---\n\n- **HR-1 — regla**\n",
                                   encoding="utf-8")
    (root / "references" / "hr-gate-matrix.json").write_text(json.dumps(
        {"rules": [{"id": "HR-1", "title": "regla", "enforcement": "gate",
                    "asset": "assets/gate/gate.py",
                    "test": "assets/gate/tests/test_gate.py"}]}), encoding="utf-8")
    (root / "assets" / "gate" / "gate.py").write_text("# gate\n", encoding="utf-8")
    (root / "assets" / "gate" / "tests" / "test_gate.py").write_text(body, encoding="utf-8")
    return root


def run_gate(skill_root, **env_extra):
    """Ejecuta el meta-gate con las marcas limpias salvo las que se pidan."""
    env = dict(os.environ)
    for key in MARKERS:
        env.pop(key, None)
    env.update({k: v for k, v in env_extra.items() if v is not None})
    return subprocess.run([sys.executable, str(GATE), str(skill_root)],
                          capture_output=True, text=True, timeout=180, env=env)


class SmokeEnvironmentTests(unittest.TestCase):
    def test_first_level_smoke_run_has_no_recursion_marker_and_depth_one(self):
        with tempfile.TemporaryDirectory() as tmp:
            probe = Path(tmp) / "probe.txt"
            skill = make_mini_skill(tmp)
            run = run_gate(skill, CI_PATTERN_NESTED="1", SMOKE_PROBE=str(probe))
            self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
            self.assertIn("HR-GATE-MATRIX OK", run.stdout)
            self.assertEqual(probe.read_text(encoding="utf-8").splitlines(),
                             ["ran depth=1"],
                             "el primer nivel cuenta profundidad 1")

    def test_test_already_running_is_not_smoke_run_again(self):
        with tempfile.TemporaryDirectory() as tmp:
            probe = Path(tmp) / "probe.txt"
            skill = make_mini_skill(tmp)
            declared = str((skill / "assets" / "gate" / "tests" / "test_gate.py").resolve())
            run = run_gate(skill, CI_PATTERN_SMOKE="1", CI_PATTERN_SMOKE_TESTS=declared,
                           SMOKE_PROBE=str(probe))
            self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
            self.assertFalse(probe.exists(),
                             "un test que ya se está ejecutando no se repite: ahí nace la recursión")
            self.assertIn("solo estructura", run.stdout)

    def test_nested_gate_still_smoke_runs_a_test_that_is_not_in_the_chain(self):
        # Un meta-gate anidado SÍ juzga los tests que no se están ejecutando:
        # si no, los fixtures que declaran tests rotos dejarían de fallar y la
        # guarda anti-recursión se convertiría en un verde silencioso.
        with tempfile.TemporaryDirectory() as tmp:
            probe = Path(tmp) / "probe.txt"
            skill = make_mini_skill(tmp)
            run = run_gate(skill, CI_PATTERN_SMOKE="1", CI_PATTERN_SMOKE_TESTS="/otro/test.py",
                           SMOKE_PROBE=str(probe))
            self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
            self.assertEqual(probe.read_text(encoding="utf-8").splitlines(),
                             ["ran depth=2"])

    def test_depth_over_the_cap_fails_closed(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_mini_skill(tmp)
            run = run_gate(skill, CI_PATTERN_SMOKE="9")
            self.assertEqual(run.returncode, 2)
            self.assertIn("profundidad", run.stderr)

    def test_malformed_depth_fails_closed(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_mini_skill(tmp)
            run = run_gate(skill, CI_PATTERN_SMOKE="dos")
            self.assertEqual(run.returncode, 2)
            self.assertIn("CI_PATTERN_SMOKE", run.stderr)


class StaticGuardsTests(unittest.TestCase):
    def test_every_subprocess_run_in_the_gate_has_a_timeout(self):
        # Regla permanente (#277): todo subprocess.run lleva timeout=. Si
        # alguien añade uno sin timeout, esta prueba falla antes de que el
        # meta-gate pueda colgarse.
        tree = ast.parse(GATE.read_text(encoding="utf-8"))
        calls = [node for node in ast.walk(tree)
                 if isinstance(node, ast.Call)
                 and isinstance(node.func, ast.Attribute)
                 and node.func.attr == "run"
                 and isinstance(node.func.value, ast.Name)
                 and node.func.value.id == "subprocess"]
        self.assertTrue(calls, "el meta-gate debe smoke-ejecutar los tests declarados")
        for call in calls:
            keywords = {kw.arg for kw in call.keywords}
            self.assertIn("timeout", keywords,
                          f"subprocess.run sin timeout= en la línea {call.lineno}")

    def test_recursion_marker_name_matches_the_cli(self):
        # La marca vive en dos ficheros porque la CLI es un script, no un
        # módulo importable: si alguien renombra una, esta prueba lo dice.
        def marker(path):
            match = re.search(r'^RECURSION_ENV = "([^"]+)"',
                              path.read_text(encoding="utf-8"), re.MULTILINE)
            self.assertIsNotNone(match, f"RECURSION_ENV no encontrada en {path}")
            return match.group(1)
        self.assertEqual(marker(GATE), marker(CLI))


if __name__ == "__main__":
    unittest.main(verbosity=2)
