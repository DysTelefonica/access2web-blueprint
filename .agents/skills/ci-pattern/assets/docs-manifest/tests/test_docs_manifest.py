#!/usr/bin/env python3
# test_docs_manifest.py — manifiesto canónico de documentos (HR-70, #719)
"""Cada documento activado existe y cada sección obligatoria existe y no
está vacía; el consumer puede añadir, nunca quitar. Solo stdlib."""

import contextlib
import importlib.machinery
import importlib.util
import io
import json
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
CLI_PATH = ASSET.parent / "bin" / "ci-pattern"


def _load(name, path, loader=None):
    spec = importlib.util.spec_from_file_location(name, path, loader=loader)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


GATE = _load("check_docs_manifest", ASSET / "check_docs_manifest.py")
CLI = _load("ci_pattern_cli_manifest", CLI_PATH,
            importlib.machinery.SourceFileLoader("ci_pattern_cli_manifest", str(CLI_PATH)))
MANIFEST = {"documents": [
    {"path": "README.md", "purpose": "Qué es el proyecto", "sections": ["Uso"],
     "required_by": "HR-48"},
    {"path": "docs/CODEBASE-GUIDE.md", "purpose": "Dónde trabajar",
     "sections": ["Páginas"], "required_by": "HR-63", "activated_by": "docs_map"},
]}


def _write(repo, rel, text):
    path = Path(repo) / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _findings(files, contract=None, manifest=MANIFEST):
    with tempfile.TemporaryDirectory() as d:
        for rel, text in files.items():
            _write(d, rel, text)
        return GATE.check(Path(d), manifest, contract or {"docs_manifest": {}})


OK = {"README.md": "# Proyecto\n\n## Uso\n\nTexto.\n"}


class DocsManifestTests(unittest.TestCase):
    def test_conforme_da_verde(self):
        self.assertEqual(_findings(OK), [])

    def test_documento_ausente_nombra_el_fichero(self):
        found = _findings({})
        self.assertTrue(any("README.md" in f and "falta el documento" in f for f in found), found)

    def test_seccion_ausente_nombra_fichero_y_seccion(self):
        found = _findings({"README.md": "# Proyecto\n"})
        self.assertTrue(any("README.md" in f and "«Uso»" in f for f in found), found)

    def test_seccion_vacia_es_roja(self):
        found = _findings({"README.md": "# P\n\n## Uso\n\n<!-- nada -->\n\n## Otra\n\ntexto\n"})
        self.assertTrue(any("«Uso»" in f and "vacía" in f for f in found), found)

    def test_titulo_dentro_de_un_bloque_cercado_no_cuenta(self):
        found = _findings({"README.md": "# P\n\n```\n## Uso\ntexto\n```\n"})
        self.assertTrue(any("«Uso»" in f for f in found), found)

    def test_bloque_canonico_embebido_cuenta_como_cuerpo_de_su_seccion(self):
        # #738: el flujo operativo generado emite, bajo «Bloque canonico», el
        # bloque gobernado con sus propios titulos de nivel 2; ese bloque ES el
        # cuerpo de la seccion que lo contiene.
        doc = ("# Flujo\n\n## 12. Bloque canónico\n\n<!-- ci-pattern-agents:begin -->\n"
               "## Flujo operativo\n\ntexto del bloque\n\n## Dónde vive\n\notro texto\n"
               "<!-- ci-pattern-agents:end -->\n")
        manifest = {"documents": [{"path": "docs/flujo.md", "purpose": "x",
                                   "sections": ["Bloque canónico"], "required_by": "HR-47"}]}
        self.assertEqual(_findings({"docs/flujo.md": doc}, manifest=manifest), [])

    def test_seccion_vacia_sin_bloque_embebido_sigue_roja(self):
        # #738: la regla solo cubre el bloque gobernado; un titulo seguido de
        # otro titulo sigue siendo una seccion vacia.
        doc = "# Flujo\n\n## Bloque canónico\n\n## Otra\n\ntexto\n"
        manifest = {"documents": [{"path": "docs/flujo.md", "purpose": "x",
                                   "sections": ["Bloque canónico"], "required_by": "HR-47"}]}
        found = _findings({"docs/flujo.md": doc}, manifest=manifest)
        self.assertTrue(any("«Bloque canónico»" in f and "vacía" in f for f in found), found)

    def test_los_titulos_del_bloque_embebido_siguen_siendo_secciones_propias(self):
        # #738: AGENTS.md exige «Flujo operativo...» y esos titulos viven dentro
        # del bloque gobernado; no dejan de contar.
        doc = "<!-- ci-pattern-agents:begin -->\n## Flujo operativo\n\ntexto\n<!-- ci-pattern-agents:end -->\n"
        manifest = {"documents": [{"path": "AGENTS.md", "purpose": "x",
                                   "sections": ["Flujo operativo"], "required_by": "HR-55"}]}
        self.assertEqual(_findings({"AGENTS.md": doc}, manifest=manifest), [])

    def test_titulo_numerado_cuenta_como_su_texto(self):
        self.assertEqual(_findings({"README.md": "## 1. Uso\n\ntexto\n"}), [])

    def test_documento_condicionado_solo_se_exige_si_su_declaracion_esta_activa(self):
        self.assertEqual(_findings(OK), [])
        found = _findings(OK, {"docs_manifest": {}, "docs_map": {}})
        self.assertTrue(any("CODEBASE-GUIDE.md" in f for f in found), found)

    def test_el_manifiesto_canonico_exige_la_arquitectura_solo_con_la_declaracion(self):
        canon = GATE.load_manifest(ASSET / "required-docs.json")
        base = {"README.md": "# P\n"}
        sin_declaracion = _findings(base, {"docs_manifest": {}}, canon)
        self.assertFalse([f for f in sin_declaracion if "architecture.md" in f or "Capas" in f])
        found = _findings(base, {"docs_manifest": {}, "architecture": {}}, canon)
        self.assertTrue(any("docs/architecture.md" in f and "HR-68" in f for f in found), found)
        doc = "# A\n\n## Árbol de módulos\n\ntexto\n"
        found = _findings({**base, "docs/architecture.md": doc}, {"docs_manifest": {}, "architecture": {}}, canon)
        self.assertTrue(any("«Capas y dependencias»" in f for f in found), found)

    def test_el_consumer_puede_anadir_documentos_y_secciones(self):
        extend = [{"path": "docs/runbook.md", "purpose": "Operación", "sections": ["Rollback"]},
                  {"path": "README.md", "purpose": "x", "sections": ["Licencia"]}]
        found = _findings(OK, {"docs_manifest": {"extend": extend}})
        self.assertTrue(any("runbook.md" in f for f in found), found)
        self.assertTrue(any("README.md" in f and "«Licencia»" in f for f in found), found)

    def test_el_consumer_no_puede_quitar_lo_del_canon(self):
        found = _findings(OK, {"docs_manifest": {"exclude": ["README.md"]}})
        self.assertTrue(any("exclude" in f for f in found), found)

    def test_manifiesto_invalido_falla_cerrado(self):
        with self.assertRaises(GATE.ManifestError):
            GATE.validate_manifest({"documents": [{"path": "README.md"}]})
        with tempfile.TemporaryDirectory() as d:
            _write(d, "m.json", "{no es json")
            with self.assertRaises(GATE.ManifestError):
                GATE.load_manifest(Path(d) / "m.json")

    def test_sin_declaracion_avisa_por_stderr_y_sale_verde(self):
        buf = io.StringIO()
        with tempfile.TemporaryDirectory() as d, contextlib.redirect_stderr(buf):
            self.assertEqual(CLI._check_docs_manifest(Path(d), {}), [])
        self.assertIn("HR-70: docs_manifest no declarado", buf.getvalue())

    def test_el_manifiesto_canonico_es_valido_y_su_flujo_operativo_no_deriva(self):
        manifest = GATE.load_manifest(ASSET / "required-docs.json")
        flow = next(d for d in manifest["documents"] if d["path"] == "docs/ci-pattern-flow.md")
        self.assertEqual(flow["sections"], [t for _k, t in CLI.OPERATING_DOC_SECTIONS])

    def test_declarado_en_el_contrato_la_cli_mide_contra_el_manifiesto_canonico(self):
        with tempfile.TemporaryDirectory() as d:
            found = CLI._check_docs_manifest(Path(d), {"docs_manifest": {}})
        self.assertTrue(any("README.md" in f and "falta el documento" in f for f in found), found)

    def test_la_plantilla_del_hub_lleva_las_secciones_del_manifiesto(self):
        manifest = GATE.load_manifest(ASSET / "required-docs.json")
        hub = next(d for d in manifest["documents"] if d["path"] == "docs/CODEBASE-GUIDE.md")
        template = (ASSET.parent / "templates" / "docs" / "CODEBASE-GUIDE.md.tmpl")
        found = GATE._sections(template.read_text(encoding="utf-8"))
        self.assertEqual([s for s in hub["sections"] if s.casefold() not in found], [])

    def test_la_plantilla_de_la_guia_de_mantenimiento_lleva_la_seccion_del_manifiesto(self):
        manifest = GATE.load_manifest(ASSET / "required-docs.json")
        page = next(d for d in manifest["documents"] if d["path"].endswith("maintainer-playbook.md"))
        template = ASSET.parent / "templates" / "docs" / "maintainer-playbook.md.tmpl"
        found = GATE._sections(template.read_text(encoding="utf-8"))
        self.assertEqual([s for s in page["sections"] if s.casefold() not in found], [])

    def test_la_plantilla_del_indice_de_decisiones_lleva_la_seccion_del_manifiesto(self):
        manifest = GATE.load_manifest(ASSET / "required-docs.json")
        page = next(d for d in manifest["documents"] if d["path"] == "docs/decisions/README.md")
        template = ASSET.parent / "templates" / "docs" / "decisions" / "README.md.tmpl"
        found = GATE._sections(template.read_text(encoding="utf-8"))
        self.assertEqual([s for s in page["sections"] if s.casefold() not in found], [])

    def test_el_esquema_del_contrato_y_el_gate_admiten_las_mismas_claves(self):
        schema = json.loads((ASSET.parent / "adoption" / "adoption-contract.schema.json")
                            .read_text(encoding="utf-8"))
        self.assertEqual(set(schema["properties"]["docs_manifest"]["properties"]),
                         set(GATE.CONTRACT_KEYS))


if __name__ == "__main__":
    unittest.main()
