#!/usr/bin/env python3
# test_pr_walkthrough.py — recorrido de PR del job requerido (HR-63, #709)
"""Un PR que crea un módulo bajo code_root sin tocar la documentación avisa
nombrando el módulo (#1093, lote 2: aviso, no bloquea); el mismo cambio con su
spoke queda limpio. Un enlace muerto en la documentación sigue bloqueando."""

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
    def test_modulo_nuevo_sin_docs_avisa_y_con_su_spoke_es_limpio(self):
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
            aviso = job(d)
            self.assertEqual(aviso.returncode, 0, aviso.stdout + aviso.stderr)
            self.assertIn("AVISO", aviso.stderr)
            self.assertIn("billing", aviso.stderr)
            write(d, "docs/codebase/map.md", "`app/users.py` usuarios\n`app/billing.py` cobros\n")
            git(d, "add", "-A")
            git(d, "commit", "-q", "--amend", "--no-edit")
            limpio = job(d)
            self.assertEqual(limpio.returncode, 0, limpio.stdout + limpio.stderr)
            self.assertNotIn("AVISO", limpio.stderr)

    def test_enlace_muerto_sigue_bloqueando_aunque_haya_modulo_sin_docs(self):
        with tempfile.TemporaryDirectory() as d:
            git(d, "init", "-q", "-b", "main")
            write(d, "app/users.py", "x = 1\n")
            write(d, "docs/codebase/map.md", "`app/users.py` usuarios\n")
            git(d, "add", "-A")
            git(d, "commit", "-q", "-m", "base")
            git(d, "checkout", "-q", "-b", "feature")
            write(d, "app/billing.py", "y = 2\n")
            write(d, "docs/codebase/map.md", "`app/users.py` usuarios\n`app/borrado.py` gone\n")
            git(d, "add", "-A")
            git(d, "commit", "-q", "-m", "enlace muerto y módulo sin docs")
            rojo = job(d)
            self.assertEqual(rojo.returncode, 1, rojo.stdout + rojo.stderr)
            self.assertIn("FAIL docs -> repo: missing app/borrado.py", rojo.stderr)
            self.assertIn("AVISO", rojo.stderr)


if __name__ == "__main__":
    unittest.main()
