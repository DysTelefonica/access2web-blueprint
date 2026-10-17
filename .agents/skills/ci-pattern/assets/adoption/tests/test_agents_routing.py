#!/usr/bin/env python3
# test_agents_routing.py — enrutado de la documentación en AGENTS.md (HR-69, #716)
"""El bloque gobernado `ci-pattern-agents` lleva la tabla «Dónde trabajar»
generada desde el contrato y el manifiesto de documentos: solo los documentos
activados, toda ruta enrutada existe y la deriva falla. Solo stdlib."""

import contextlib
import importlib.machinery
import importlib.util
import io
import json
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parents[2]
CLI_PATH = ASSET / "bin" / "ci-pattern"


def _load(name, path, loader=None):
    spec = importlib.util.spec_from_file_location(name, path, loader=loader)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


GATE = _load("check_docs_manifest", ASSET / "docs-manifest" / "check_docs_manifest.py")
CLI = _load("ci_pattern_cli_routing", CLI_PATH,
            importlib.machinery.SourceFileLoader("ci_pattern_cli_routing", str(CLI_PATH)))
ALL = {"docs_manifest": {}, "docs_map": {}, "docs_decisions": {}, "architecture": {}}


def repo_with(root, contract, files=("docs/CODEBASE-GUIDE.md", "docs/codebase/maintainer-playbook.md",
                                      "docs/decisions/README.md", "docs/architecture.md"), agents=True):
    repo = Path(root)
    (repo / ".github").mkdir()
    (repo / ".github" / "ci-pattern-adoption.json").write_text(json.dumps(contract), encoding="utf-8")
    for rel in files:
        (repo / rel).parent.mkdir(parents=True, exist_ok=True)
        (repo / rel).write_text("# x\n", encoding="utf-8")
    if agents:
        (repo / "AGENTS.md").write_text("# Proyecto\n\n" + CLI._agents_block_for(repo) + "\n", encoding="utf-8")
    return repo


class AgentsRoutingTests(unittest.TestCase):
    def test_sin_declaracion_el_bloque_es_byte_a_byte_el_actual(self):
        with tempfile.TemporaryDirectory() as d:
            repo = repo_with(d, {}, agents=False)
            self.assertEqual(CLI._agents_block_for(repo), CLI.OPERATING_DOC_AGENTS_BLOCK)

    def test_solo_se_enrutan_los_documentos_activados(self):
        with tempfile.TemporaryDirectory() as d:
            block = CLI._agents_block_for(repo_with(d, {"docs_manifest": {}, "docs_map": {}}, agents=False))
        self.assertIn("## Dónde trabajar", block)
        self.assertIn("docs/CODEBASE-GUIDE.md", block)
        self.assertIn("maintainer-playbook.md", block)
        self.assertNotIn("docs/decisions", block)

    def test_con_decisiones_activadas_se_enruta_el_registro(self):
        with tempfile.TemporaryDirectory() as d:
            block = CLI._agents_block_for(repo_with(d, ALL, agents=False))
        self.assertIn("docs/decisions/", block)

    def test_todo_documento_activado_con_ruta_tiene_su_fila(self):
        manifest = GATE.load_manifest(ASSET / "docs-manifest" / "required-docs.json")
        routed = {path for _w, _r, path in GATE.routing_rows(manifest, ALL)}
        expected = {d["path"] for d in manifest["documents"] if d.get("route")}
        self.assertEqual(routed, expected)
        self.assertTrue(expected)

    def test_una_ruta_enrutada_que_no_existe_es_roja_y_nombra_la_ruta(self):
        with tempfile.TemporaryDirectory() as d:
            repo = repo_with(d, ALL, files=("docs/CODEBASE-GUIDE.md",))
            found = CLI._check_agents_block(repo)
        self.assertTrue(any("docs/decisions/README.md" in f and "no existe" in f for f in found), found)

    def test_conforme_da_verde(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(CLI._check_agents_block(repo_with(d, ALL)), [])

    def test_la_deriva_entre_agents_y_la_regeneracion_falla(self):
        with tempfile.TemporaryDirectory() as d:
            repo = repo_with(d, ALL)
            text = (repo / "AGENTS.md").read_text(encoding="utf-8")
            (repo / "AGENTS.md").write_text(text.replace("docs/decisions/", "docs/otra/"), encoding="utf-8")
            found = CLI._check_agents_block(repo)
        self.assertTrue(any("alterado" in f or "Dónde trabajar" in f for f in found), found)

    def test_el_subcomando_imprime_el_bloque_regenerado(self):
        with tempfile.TemporaryDirectory() as d:
            repo = repo_with(d, ALL, agents=False)
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                code = CLI.cmd_adoption(["render-agents-block", str(repo)])
            self.assertEqual((code, buf.getvalue().rstrip("\n")), (0, CLI._agents_block_for(repo)))


if __name__ == "__main__":
    unittest.main()
