#!/usr/bin/env python3
"""HR-64 (#710): evidencia de documentación por PR. stdlib exclusivamente.

Un PR que toca código cubierto por una página del mapa toca la página o lleva
la sección «Documentación» no vacía; las checklists por tipo de cambio existen
y el cuerpo del PR enlaza la de su tipo. Sin declaración avisa y sale verde.

#1093 R1 (decisión del dueño, 2026-10-23): HR-64 es AVISO — el gate lo
imprime y sale 0; el PR no se bloquea por la evidencia de documentación."""

import contextlib
import importlib.util
import io
import re
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
        advisories = []
        changed = subprocess.run(["git", "-C", str(repo), "diff", "--name-only", f"{base}..{sha(repo)}"],
                                 check=True, capture_output=True, text=True).stdout.split()
        MOD.check_docs_evidence(pr, repo, changed, {"docs_evidence": evidence}, advisories)
        return advisories


class DocsEvidenceTests(unittest.TestCase):
    def test_codigo_y_pagina_en_el_mismo_diff_es_verde(self):
        self.assertEqual(run({"app/users/a.py": "x = 2\n", "docs/codebase/users.md": "nuevo\n"}), [])

    def test_solo_codigo_con_seccion_documentacion_es_verde(self):
        body = LINK + "\n## Documentación\n\nSolo renombra una variable local; la página no cambia.\n"
        self.assertEqual(run({"app/users/a.py": "x = 2\n"}, body=body), [])

    def test_solo_codigo_sin_seccion_es_aviso_y_nombra_pagina_y_fichero(self):
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

    def test_checklist_inexistente_o_ancla_inexistente_es_aviso(self):
        bad = {**EVIDENCE, "checklists": {"type:feature": "docs/codebase/ghost.md#feature"}}
        found = run({"otro/b.py": "y\n"}, evidence=bad, body="Checklist: docs/codebase/ghost.md#feature\n")
        self.assertTrue(any("ghost.md" in f for f in found), found)
        bad = {**EVIDENCE, "checklists": {"type:feature": PLAYBOOK + "#bug"}}
        found = run({"otro/b.py": "y\n"}, evidence=bad, body=f"Checklist: {PLAYBOOK}#bug\n")
        self.assertTrue(any("#bug" in f for f in found), found)

    def test_el_cuerpo_debe_enlazar_la_checklist_de_su_tipo(self):
        found = run({"otro/b.py": "y\n"}, body="sin enlace\n")
        self.assertTrue(any("type:feature" in f and "maintainer-playbook" in f for f in found), found)

    # #772: la plantilla de PR del canon trae los enlaces a las checklists por tipo y
    # cada uno resuelve a un encabezado de la guia de mantenimiento del canon.
    ROOT = Path(__file__).resolve().parents[6]
    TEMPLATE = ROOT / ".github" / "PULL_REQUEST_TEMPLATE.md"
    PLAYBOOK = Path(__file__).resolve().parents[2] / "templates" / "docs" / "maintainer-playbook.md.tmpl"
    TYPES = ("feature", "bug", "refactor", "docs", "chore")

    def test_template_links_one_checklist_per_type(self):
        text = self.TEMPLATE.read_text(encoding="utf-8")
        for change_type in self.TYPES:
            self.assertIn(f"{PLAYBOOK}#{change_type}", text, change_type)

    def test_every_linked_anchor_is_a_playbook_heading(self):
        slugs = {re.sub(r"[^\w\- ]", "", m.group(1)).strip().lower().replace(" ", "-")
                 for m in re.finditer(r"^#{1,6}\s+(.*?)\s*#*$",
                                      self.PLAYBOOK.read_text(encoding="utf-8"), re.M)}
        for change_type in self.TYPES:
            self.assertIn(change_type, slugs)

    def test_links_are_the_ones_the_policy_example_expects(self):
        text = self.TEMPLATE.read_text(encoding="utf-8")
        links = re.findall(r"\(([^)]*maintainer-playbook\.md#[a-z]+)\)", text)
        self.assertEqual(sorted(links), sorted(f"{PLAYBOOK}#{t}" for t in self.TYPES))


if __name__ == "__main__":
    unittest.main()
