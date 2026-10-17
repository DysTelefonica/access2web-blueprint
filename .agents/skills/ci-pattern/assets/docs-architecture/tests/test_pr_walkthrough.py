#!/usr/bin/env python3
# test_pr_walkthrough.py — recorrido de PR del job requerido (HR-68, #714)
"""Un PR que crea un módulo bajo code_root sin tocar docs/architecture.md deja el
job rojo nombrando el módulo; con su fila en el árbol pero sin capa sigue rojo;
con su fila, su capa y la configuración del linter regenerada queda verde. Usa
import-linter REAL (instalado en CI; si falta, el job sale rojo, no se salta)."""

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parent.parent / "check_docs_architecture.py"
BLOCK = {"layers": [{"name": "web", "modules": ["app/routes"]},
                    {"name": "core", "modules": ["app/domain"]}],
         "allowed": {"web": ["core"], "core": []}}


def git(repo, *args):
    subprocess.run(["git", "-c", "user.name=t", "-c", "user.email=t@e.invalid",
                    "-c", "commit.gpgsign=false", *args], cwd=repo, check=True,
                   capture_output=True)


def write(repo, rel, text):
    path = Path(repo) / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def doc(tree, block):
    rows = "".join(f"| `{p}` | hace algo |\n" for p in tree)
    return ("# Arquitectura\n\n## Árbol de módulos\n\n| Módulo | Responsabilidad |\n|---|---|\n"
            f"{rows}\n## Capas y dependencias\n\n```architecture\n{json.dumps(block)}\n```\n")


def job(repo, *extra):
    """El job requerido: el gate sobre el árbol del PR, con la configuración del linter."""
    return subprocess.run([sys.executable, str(GATE), str(repo), "--code-root", "app",
                           "--code-glob", "*.py", "--linter-config", ".importlinter", *extra],
                          capture_output=True, text=True)


def regenerate(repo):
    return subprocess.run([sys.executable, str(GATE), str(repo), "--code-root", "app",
                           "--code-glob", "*.py", "--linter-config", ".importlinter",
                           "--write-config"], capture_output=True, text=True)


class PrWalkthroughTests(unittest.TestCase):
    def test_modulo_nuevo_sin_docs_es_rojo_sin_capa_es_rojo_y_con_fila_y_capa_es_verde(self):
        with tempfile.TemporaryDirectory() as d:
            git(d, "init", "-q", "-b", "main")
            for name in ("routes", "domain"):
                write(d, f"app/{name}/__init__.py", "")
                write(d, f"app/{name}/x.py", "x = 1\n")
            write(d, "app/__init__.py", "")
            write(d, "docs/architecture.md", doc(["app/routes", "app/domain"], BLOCK))
            self.assertEqual(regenerate(d).returncode, 0)
            git(d, "add", "-A")
            git(d, "commit", "-q", "-m", "base")
            self.assertEqual(job(d).returncode, 0, "la base debe estar al día")

            git(d, "checkout", "-q", "-b", "feature")
            write(d, "app/billing/__init__.py", "")
            write(d, "app/billing/x.py", "y = 2\n")
            git(d, "add", "-A")
            git(d, "commit", "-q", "-m", "nuevo módulo sin docs")
            red = job(d)
            self.assertEqual(red.returncode, 1)
            self.assertIn("app/billing", red.stderr)

            write(d, "docs/architecture.md",
                  doc(["app/routes", "app/domain", "app/billing"], BLOCK))
            git(d, "add", "-A")
            git(d, "commit", "-q", "--amend", "--no-edit")
            no_layer = job(d)
            self.assertEqual(no_layer.returncode, 1)
            self.assertIn("ninguna capa", no_layer.stderr)

            block = {**BLOCK, "layers": [BLOCK["layers"][0],
                                         {"name": "core", "modules": ["app/domain", "app/billing"]}]}
            write(d, "docs/architecture.md", doc(["app/routes", "app/domain", "app/billing"], block))
            self.assertEqual(job(d).returncode, 1, "el linter no refleja aún la capa")
            self.assertEqual(regenerate(d).returncode, 0)
            git(d, "add", "-A")
            git(d, "commit", "-q", "--amend", "--no-edit")
            self.assertEqual(job(d).returncode, 0)


if __name__ == "__main__":
    unittest.main()
