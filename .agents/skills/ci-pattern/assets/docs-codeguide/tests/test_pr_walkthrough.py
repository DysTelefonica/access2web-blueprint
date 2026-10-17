#!/usr/bin/env python3
# test_pr_walkthrough.py — recorrido de PR del job requerido (HR-63, #709)
"""Un PR que crea un módulo bajo code_root sin tocar la documentación deja el
job rojo nombrando el módulo; el mismo cambio con su spoke lo deja verde."""

import subprocess
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parent.parent / "check_docs_codeguide.py"


def git(repo, *args):
    subprocess.run(["git", "-c", "user.name=t", "-c", "user.email=t@e.invalid",
                    "-c", "commit.gpgsign=false", *args], cwd=repo, check=True,
                   capture_output=True)


def write(repo, rel, text):
    path = Path(repo) / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def job(repo):
    """El job requerido: el gate sobre el árbol del PR."""
    return subprocess.run(["python3", str(GATE), str(repo), "--docs-path", "docs/codebase",
                           "--code-root", "app", "--code-glob", "*.py"],
                          capture_output=True, text=True)


class PrWalkthroughTests(unittest.TestCase):
    def test_modulo_nuevo_sin_docs_es_rojo_y_con_su_spoke_es_verde(self):
        with tempfile.TemporaryDirectory() as d:
            git(d, "init", "-q", "-b", "main")
            write(d, "app/users.py", "x = 1\n")
            write(d, "docs/codebase/map.md", "`app/users.py` usuarios\n")
            git(d, "add", "-A")
            git(d, "commit", "-q", "-m", "base")
            self.assertEqual(job(d).returncode, 0, "la base debe estar al día")
            git(d, "checkout", "-q", "-b", "feature")
            write(d, "app/billing.py", "y = 2\n")
            git(d, "add", "-A")
            git(d, "commit", "-q", "-m", "nuevo módulo sin docs")
            red = job(d)
            self.assertEqual(red.returncode, 1)
            self.assertIn("billing", red.stderr)
            write(d, "docs/codebase/map.md", "`app/users.py` usuarios\n`app/billing.py` cobros\n")
            git(d, "add", "-A")
            git(d, "commit", "-q", "--amend", "--no-edit")
            self.assertEqual(job(d).returncode, 0)


if __name__ == "__main__":
    unittest.main()
