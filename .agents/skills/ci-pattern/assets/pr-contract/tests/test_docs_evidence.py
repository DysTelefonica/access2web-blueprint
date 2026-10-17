#!/usr/bin/env python3
"""HR-64 (#710): evidencia de documentación por PR. stdlib exclusivamente.

Un PR que toca código cubierto por una página del mapa toca la página o lleva
la sección «Documentación» no vacía; las checklists por tipo de cambio existen
y el cuerpo del PR enlaza la de su tipo. Sin declaración avisa y sale verde."""

import contextlib
import importlib.util
import io
import subprocess
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_pr_contract.py"
spec = importlib.util.spec_from_file_location("check_pr_contract_docs", GATE)
MOD = importlib.util.module_from_spec(spec)
spec.loader.exec_module(MOD)

PLAYBOOK = "docs/codebase/maintainer-playbook.md"
EVIDENCE = {"pages": [{"page": "docs/codebase/users.md", "covers": ["app/users/*"]}],
            "checklists": {"type:feature": PLAYBOOK + "#feature"}}
LINK = f"Checklist: [{PLAYBOOK}#feature]({PLAYBOOK}#feature)\n"


def git(repo, *args):
    subprocess.run(["git", "-c", "user.name=t", "-c", "user.email=t@e.invalid",
                    "-c", "commit.gpgsign=false", "-C", str(repo), *args],
                   check=True, capture_output=True)


def sha(repo):
    return subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"], check=True,
                          capture_output=True, text=True).stdout.strip()


def run(files, body=LINK, evidence=EVIDENCE, labels=("type:feature",)):
    """Base con el mapa al día + un commit con `files` (ruta -> texto)."""
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d)
        git(repo, "init", "-q", "-b", "main")
        for rel, text in {"app/users/a.py": "x = 1\n", "docs/codebase/users.md": "usuarios\n",
                          PLAYBOOK: "# Checklists\n\n## Feature\n\n- [ ] docs\n"}.items():
            (repo / rel).parent.mkdir(parents=True, exist_ok=True)
            (repo / rel).write_text(text, encoding="utf-8")
        git(repo, "add", "-A")
        git(repo, "commit", "-q", "-m", "base")
        base = sha(repo)
        for rel, text in files.items():
            (repo / rel).parent.mkdir(parents=True, exist_ok=True)
            (repo / rel).write_text(text, encoding="utf-8")
        git(repo, "add", "-A")
        git(repo, "commit", "-q", "-m", "cambio")
        pr = {"body": body, "labels": [{"name": n} for n in labels]}
        findings = []
        changed = subprocess.run(["git", "-C", str(repo), "diff", "--name-only", f"{base}..{sha(repo)}"],
                                 check=True, capture_output=True, text=True).stdout.split()
        MOD.check_docs_evidence(pr, repo, changed, {"docs_evidence": evidence}, findings)
        return findings


class DocsEvidenceTests(unittest.TestCase):
    def test_codigo_y_pagina_en_el_mismo_diff_es_verde(self):
        self.assertEqual(run({"app/users/a.py": "x = 2\n", "docs/codebase/users.md": "nuevo\n"}), [])

    def test_solo_codigo_con_seccion_documentacion_es_verde(self):
        body = LINK + "\n## Documentación\n\nSolo renombra una variable local; la página no cambia.\n"
        self.assertEqual(run({"app/users/a.py": "x = 2\n"}, body=body), [])

    def test_solo_codigo_sin_seccion_es_rojo_y_nombra_pagina_y_fichero(self):
        found = run({"app/users/a.py": "x = 2\n"})
        self.assertTrue(any("docs/codebase/users.md" in f and "app/users/a.py" in f for f in found), found)

    def test_seccion_vacia_o_solo_comentario_no_exonera(self):
        body = LINK + "\n## Documentación\n\n<!-- motivo -->\n"
        self.assertTrue(run({"app/users/a.py": "x = 2\n"}, body=body))

    def test_codigo_no_cubierto_por_ninguna_pagina_es_verde(self):
        self.assertEqual(run({"otro/b.py": "y = 1\n"}), [])

    def test_sin_declaracion_avisa_por_stderr_y_no_falla(self):
        buf, findings = io.StringIO(), []
        with contextlib.redirect_stderr(buf):
            MOD.check_docs_evidence({"body": ""}, Path("."), [], {}, findings)
        self.assertEqual(findings, [])
        self.assertIn("HR-64: docs_evidence no declarado", buf.getvalue())

    def test_declaracion_invalida_falla_cerrado(self):
        with self.assertRaises(MOD.DocsEvidenceConfigError):
            run({"app/users/a.py": "x\n"}, evidence={"pages": []})

    def test_checklist_inexistente_o_ancla_inexistente_es_roja(self):
        bad = {**EVIDENCE, "checklists": {"type:feature": "docs/codebase/ghost.md#feature"}}
        found = run({"otro/b.py": "y\n"}, evidence=bad, body="Checklist: docs/codebase/ghost.md#feature\n")
        self.assertTrue(any("ghost.md" in f for f in found), found)
        bad = {**EVIDENCE, "checklists": {"type:feature": PLAYBOOK + "#bug"}}
        found = run({"otro/b.py": "y\n"}, evidence=bad, body=f"Checklist: {PLAYBOOK}#bug\n")
        self.assertTrue(any("#bug" in f for f in found), found)

    def test_el_cuerpo_debe_enlazar_la_checklist_de_su_tipo(self):
        found = run({"otro/b.py": "y\n"}, body="sin enlace\n")
        self.assertTrue(any("type:feature" in f and "maintainer-playbook" in f for f in found), found)


if __name__ == "__main__":
    unittest.main()
