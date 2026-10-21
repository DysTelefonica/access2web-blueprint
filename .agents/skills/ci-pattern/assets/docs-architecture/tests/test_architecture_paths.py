#!/usr/bin/env python3
# test_architecture_paths.py — capas por patrón de ruta del documento de arquitectura (HR-68, #751)
"""Bloque v2: patrones de ruta con primera coincidencia, ``exempt``, expansión a
paquetes y contratos de import-linter REAL (instalado en CI; si falta, falla)."""

import importlib.util
import json
import shutil
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parents[1]


def _load(name):
    spec = importlib.util.spec_from_file_location(name, ASSET / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


paths = _load("architecture_paths")
imports = _load("architecture_imports")
gate = _load("check_docs_architecture")

ROOTS = [("app", "*.py")]
BLOCK = {"layers": [{"name": "domain", "paths": ["app/modules/*/domain/**"], "allowed": ["domain"]},
                    {"name": "infrastructure", "paths": ["app/core/*.py"], "allowed": ["domain", "infrastructure"]},
                    {"name": "delivery", "paths": ["app/modules/**"],
                     "allowed": ["domain", "infrastructure", "delivery"]}],
         "exempt": ["app/__init__.py", "app/core/__init__.py", "app/modules/__init__.py", "app/static/**"]}
TREE = ("| Módulo | Responsabilidad |\n|---|---|\n| `app/core/` | ajustes |\n"
        "| `app/modules/animals/` | animales |\n")


def _doc(block=BLOCK, marks=False):
    body = json.dumps(block)
    fenced = (f"<!-- architecture:begin -->\n```json\n{body}\n```\n<!-- architecture:end -->\n" if marks
              else f"```architecture\n{body}\n```\n")
    return f"# Arquitectura\n\n## Árbol de módulos\n\n### Raíz\n\n{TREE}\n## Capas y dependencias\n\n{fenced}"


def _repo(root, domain_source="x = 1\n", extra=None):
    repo = Path(root)
    files = {"app/__init__.py": "", "app/core/__init__.py": "", "app/core/config.py": "y = 1\n",
             "app/modules/__init__.py": "", "app/modules/animals/__init__.py": "",
             "app/modules/animals/routes.py": "from app.modules.animals.domain import animal\n",
             "app/modules/animals/domain/__init__.py": "", "app/modules/animals/domain/animal.py": domain_source,
             "app/static/site.css": "a{}", **(extra or {})}
    for rel, text in files.items():
        (repo / rel).parent.mkdir(parents=True, exist_ok=True)
        (repo / rel).write_text(text, encoding="utf-8")
    return repo


def _normalized(repo, block=BLOCK):
    return paths.normalize(repo, block, ROOTS)


class ExpansionTests(unittest.TestCase):
    def test_el_primer_patron_decide_y_las_rutas_se_compactan_a_paquetes(self):
        with tempfile.TemporaryDirectory() as d:
            data, findings = _normalized(_repo(d))
        modules = {layer["name"]: layer["modules"] for layer in data["layers"]}
        self.assertEqual(findings, [])
        self.assertEqual(modules["domain"], ["app/modules/animals/domain"])
        self.assertEqual(modules["infrastructure"], ["app/core"])
        self.assertEqual(modules["delivery"], ["app/modules/animals/__init__.py", "app/modules/animals/routes.py"])

    def test_allowed_dentro_de_la_capa_se_eleva_al_mapa_del_bloque(self):
        with tempfile.TemporaryDirectory() as d:
            data, _ = _normalized(_repo(d))
        self.assertEqual(data["allowed"]["domain"], ["domain"])
        self.assertEqual(data["allowed"]["delivery"], ["domain", "infrastructure", "delivery"])

    def test_un_fichero_fuente_sin_capa_ni_exempt_es_rojo_nombrandolo(self):
        with tempfile.TemporaryDirectory() as d:
            _, findings = _normalized(_repo(d, extra={"app/extra.py": "z = 1\n"}))
        self.assertEqual(findings, ["el fichero app/extra.py no pertenece a ninguna capa ni a «exempt»"])

    def test_un_patron_que_no_casa_nada_es_rojo(self):
        block = {**BLOCK, "layers": BLOCK["layers"] + [{"name": "ghost", "paths": ["app/ghost/**"], "allowed": []}]}
        with tempfile.TemporaryDirectory() as d:
            _, findings = _normalized(_repo(d), block)
        self.assertEqual(findings, ["la capa ghost declara el patrón app/ghost/** y no casa ningún fichero"])

    def test_un_exempt_que_no_casa_nada_es_rojo(self):
        block = {**BLOCK, "exempt": BLOCK["exempt"] + ["app/nada/**"]}
        with tempfile.TemporaryDirectory() as d:
            _, findings = _normalized(_repo(d), block)
        self.assertEqual(findings, ["«exempt» declara el patrón app/nada/** y no casa ningún fichero"])

    def test_exempt_gana_a_las_capas(self):
        with tempfile.TemporaryDirectory() as d:
            data, findings = _normalized(_repo(d))
        self.assertEqual(findings, [])
        owned = [m for layer in data["layers"] for m in layer["modules"]]
        self.assertNotIn("app/core/__init__.py", owned)

    def test_una_capa_mal_formada_es_roja(self):
        block = {"layers": [{"name": "x"}]}
        with tempfile.TemporaryDirectory() as d:
            _, findings = _normalized(_repo(d), block)
        self.assertTrue(any("capa mal formada" in f for f in findings), findings)


class OverridesTests(unittest.TestCase):
    """``overrides``: una ruta exacta decide su capa antes que los patrones (#805)."""

    OVERRIDDEN = {**BLOCK, "overrides": [{"path": "app/core/config.py", "layer": "delivery"}]}

    def test_un_override_manda_el_fichero_a_su_capa_aunque_un_patron_lo_reclame(self):
        with tempfile.TemporaryDirectory() as d:
            data, findings = _normalized(_repo(d), self.OVERRIDDEN)
        modules = {layer["name"]: layer["modules"] for layer in data["layers"]}
        self.assertEqual(findings, [])
        self.assertIn("app/core", modules["delivery"])
        self.assertEqual(modules["infrastructure"], [])

    def test_sin_overrides_el_resultado_no_cambia(self):
        with tempfile.TemporaryDirectory() as d:
            data, _ = _normalized(_repo(d))
        modules = {layer["name"]: layer["modules"] for layer in data["layers"]}
        self.assertEqual(modules["infrastructure"], ["app/core"])

    def test_exempt_sigue_ganando_a_un_override(self):
        block = {**BLOCK, "overrides": [{"path": "app/core/__init__.py", "layer": "delivery"}]}
        with tempfile.TemporaryDirectory() as d:
            data, findings = _normalized(_repo(d), block)
        self.assertEqual(findings, [])
        self.assertNotIn("app/core/__init__.py", [m for layer in data["layers"] for m in layer["modules"]])

    def test_un_override_a_un_fichero_inexistente_es_rojo(self):
        block = {**BLOCK, "overrides": [{"path": "app/core/ghost.py", "layer": "delivery"}]}
        with tempfile.TemporaryDirectory() as d:
            _, findings = _normalized(_repo(d), block)
        self.assertEqual(findings, ["«overrides» declara la ruta app/core/ghost.py y no casa ningún fichero"])

    def test_un_override_a_una_capa_no_declarada_es_rojo(self):
        block = {**BLOCK, "overrides": [{"path": "app/core/config.py", "layer": "nada"}]}
        with tempfile.TemporaryDirectory() as d:
            _, findings = _normalized(_repo(d), block)
        self.assertEqual(findings, ["«overrides» manda app/core/config.py a la capa nada, que no está declarada"])

    def test_un_override_mal_formado_es_rojo(self):
        block = {**BLOCK, "overrides": [{"path": "app/core/config.py"}]}
        with tempfile.TemporaryDirectory() as d:
            _, findings = _normalized(_repo(d), block)
        self.assertTrue(any("override mal formado" in f for f in findings), findings)


class DocumentTests(unittest.TestCase):
    def _findings(self, repo, doc, linter_config=None):
        (repo / "docs").mkdir(exist_ok=True)
        (repo / "docs" / "architecture.md").write_text(doc, encoding="utf-8")
        return gate.check_architecture(repo, "docs/architecture.md", ROOTS, linter_config)

    def test_arbol_bajo_subtitulos_con_barra_final_y_descendientes_cuadra(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(self._findings(_repo(d), _doc()), [])

    def test_las_marcas_begin_end_con_fence_json_se_aceptan(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(self._findings(_repo(d), _doc(marks=True)), [])

    def test_un_modulo_nuevo_sin_fila_ni_descendientes_es_rojo(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, extra={"app/billing/__init__.py": "", "app/billing/x.py": "b = 1\n"})
            found = self._findings(repo, _doc())
        self.assertTrue(any("app/billing" in f and "código → árbol" in f for f in found), found)

    def test_una_fila_que_apunta_a_un_camino_inexistente_es_roja(self):
        with tempfile.TemporaryDirectory() as d:
            found = self._findings(_repo(d), _doc().replace("`app/core/`", "`app/ghost/`"))
        self.assertTrue(any("app/ghost" in f and "no existe" in f for f in found), found)


class LinterTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(shutil.which("lint-imports"),
                             "import-linter no está instalado: falla cerrado, no se salta (HR-3)")

    def _findings(self, repo):
        (repo / "docs").mkdir(exist_ok=True)
        (repo / "docs" / "architecture.md").write_text(_doc(), encoding="utf-8")
        self.assertEqual(gate.write_config(repo, "docs/architecture.md", ".importlinter", roots=ROOTS), [])
        return gate.check_architecture(repo, "docs/architecture.md", ROOTS, ".importlinter")

    def test_el_hexagono_interno_de_un_modulo_se_impone_sin_falsos_positivos(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(self._findings(_repo(d)), [])

    def test_un_import_prohibido_dentro_de_un_modulo_nombra_fichero_linea_y_regla(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, domain_source="\nfrom app.core import config\n")
            found = self._findings(repo)
        self.assertEqual(found, ["app/modules/animals/domain/animal.py:2: import prohibido por la regla "
                                 "«domain must not import infrastructure»"])


if __name__ == "__main__":
    unittest.main()
