#!/usr/bin/env python3
# test_architecture_imports.py — reglas de imports del documento de arquitectura (HR-68, #714)
"""Config derivada del bloque, divergencia, linter ausente y violaciones con
import-linter REAL (instalado en CI; si falta, el test falla, no se salta)."""

import importlib.util
import json
import os
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


imports = _load("architecture_imports")
gate = _load("check_docs_architecture")

BLOCK = {"layers": [{"name": "web", "modules": ["app/routes"]},
                    {"name": "core", "modules": ["app/domain"]}],
         "allowed": {"web": ["core"], "core": []}}
DOC = ("# Arquitectura\n\n## Árbol de módulos\n\n| Módulo | Responsabilidad |\n|---|---|\n"
       "| `app/routes` | HTTP |\n| `app/domain` | reglas |\n\n## Capas y dependencias\n\n"
       "```architecture\n" + json.dumps(BLOCK) + "\n```\n")
ROOTS = [("app", "*.py")]


def _repo(root, domain_source="x = 1\n", config=True):
    repo = Path(root)
    files = {"app/__init__.py": "", "app/routes/__init__.py": "",
             "app/routes/views.py": "from app.domain import models\n",
             "app/domain/__init__.py": "", "app/domain/models.py": domain_source,
             "docs/architecture.md": DOC}
    for rel, text in files.items():
        (repo / rel).parent.mkdir(parents=True, exist_ok=True)
        (repo / rel).write_text(text, encoding="utf-8")
    if config:
        (repo / ".importlinter").write_text(imports.render_importlinter(BLOCK), encoding="utf-8")
    return repo


def _findings(repo):
    return gate.check_architecture(repo, "docs/architecture.md", ROOTS, ".importlinter")


class RenderTests(unittest.TestCase):
    def test_solo_los_pares_no_permitidos_son_contratos_forbidden(self):
        text = imports.render_importlinter(BLOCK)
        self.assertEqual(imports.forbidden_pairs(BLOCK), [("core", "web")])
        self.assertIn("[importlinter:contract:core-must-not-import-web]", text)
        self.assertIn("source_modules =\n    app.domain\nforbidden_modules =\n    app.routes", text)
        self.assertNotIn("web-must-not-import-core", text)

    def test_ficheros_planos_y_raices_unicas(self):
        block = {"layers": [{"name": "a", "modules": ["app/main.py"]},
                            {"name": "b", "modules": ["lib/x"]}], "allowed": {"a": [], "b": []}}
        text = imports.render_importlinter(block)
        self.assertIn("root_packages =\n    app\n    lib", text)
        self.assertIn("    app.main\n", text)

    def test_la_salida_es_determinista(self):
        self.assertEqual(imports.render_importlinter(BLOCK), imports.render_importlinter(BLOCK))


class GateTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(shutil.which("lint-imports"),
                             "import-linter no está instalado: falla cerrado, no se salta (HR-3)")

    def test_codigo_conforme_da_verde(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(_findings(_repo(d)), [])

    def test_import_prohibido_nombra_fichero_linea_y_regla(self):
        with tempfile.TemporaryDirectory() as d:
            found = _findings(_repo(d, domain_source="\nfrom app.routes import views\n"))
        self.assertEqual(found, ["app/domain/models.py:2: import prohibido por la regla "
                                 "«core must not import web»"])

    def test_regla_del_documento_sin_reflejo_en_el_linter_es_roja(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d)
            stale = (repo / ".importlinter").read_text(encoding="utf-8")
            block = {**BLOCK, "allowed": {"web": ["core"], "core": ["web"]}}
            (repo / "docs" / "architecture.md").write_text(
                DOC.replace(json.dumps(BLOCK), json.dumps(block)), encoding="utf-8")
            found = _findings(repo)
            self.assertTrue(any(".importlinter" in f and "no coincide" in f for f in found), found)
            self.assertEqual((repo / ".importlinter").read_text(encoding="utf-8"), stale)

    def test_configuracion_ausente_es_roja(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertTrue(any("falta la configuración" in f for f in _findings(_repo(d, config=False))))

    def test_write_config_regenera_la_configuracion_desde_el_documento(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, config=False)
            self.assertEqual(gate.write_config(repo, "docs/architecture.md", ".importlinter"), [])
            self.assertEqual(_findings(repo), [])


HEADER = ("one must not import two\n-----------------------\n\n"
          "app.one is not allowed to import \napp.two:\n\n")
WIDE = HEADER + "-   app.one.delivery.http.admin_routes -> app.two.x (l.1)\n\n"
# Salida real de import-linter 2.15 a 80 columnas: se parte tras «-> » y sigue sin sangría.
WRAPPED = HEADER + "-   app.one.delivery.http.admin_routes -> \napp.two.x (l.1)\n\n"
EXPECTED = ["app/one/delivery/http/admin_routes:1: import prohibido por la regla «one must not import two»"]


class WrappedOutputTests(unittest.TestCase):
    def parse(self, output):
        with tempfile.TemporaryDirectory() as d:
            return imports.parse_violations(Path(d), output)

    def test_la_salida_partida_tras_la_flecha_se_analiza_igual_que_la_ancha(self):
        self.assertEqual(self.parse(WIDE), EXPECTED)
        self.assertEqual(self.parse(WRAPPED), EXPECTED)

    def test_la_salida_partida_antes_de_la_linea_se_analiza(self):
        out = HEADER + "-   app.one.delivery.http.admin_routes -> app.two.x \n(l.1)\n\n"
        self.assertEqual(self.parse(out), EXPECTED)

    def test_una_ruta_partida_en_mitad_de_la_palabra_se_une_sin_espacio(self):
        out = HEADER + "-   app.one.delivery.http.admin_routes -> app.t\nwo.x (l.1)\n\n"
        self.assertEqual(self.parse(out), EXPECTED)

    def test_el_texto_que_sigue_a_un_registro_completo_no_se_le_pega(self):
        out = WIDE + "Analyzed 3 files\n-   app.one.b -> app.two.y (l.4)\n"
        found = self.parse(out)
        self.assertEqual(len(found), 2)
        self.assertTrue(found[1].startswith("app/one/b:4"))

    def test_el_adaptador_ejecuta_el_linter_con_columns_ancho(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d)
            stub = Path(d) / "stub-linter"
            body = ("#!/bin/sh\nif [ \"${COLUMNS:-80}\" -ge 200 ]; then\n"
                    "printf '%s' '" + WIDE.replace("'", "") + "'\nelse\n"
                    "printf '%s' '" + HEADER + "-   app.one.delivery.http.admin_r\noutes -> app.two.x (l.1)\n'\nfi\nexit 1\n")
            stub.write_text(body, encoding="utf-8")
            stub.chmod(0o755)
            found = imports.check_imports(repo, BLOCK, ".importlinter", command=str(stub))
        self.assertEqual(found, EXPECTED)


def _doc_for(block, tree):
    rows = "".join(f"| `{m}` | x |\n" for m in tree)
    return ("# Arquitectura\n\n## Árbol de módulos\n\n| Módulo | Responsabilidad |\n|---|---|\n" + rows +
            "\n## Capas y dependencias\n\n```architecture\n" + json.dumps(block) + "\n```\n")


# Dos aplicaciones (lan, exp) con las mismas capas: la forma de access2web-blueprint.
TREE2 = ["app/lan/domain", "app/lan/delivery", "app/lan/di", "app/exp/domain"]
BLOCK2 = {"layers": [{"name": "domain", "modules": ["app/lan/domain", "app/exp/domain"]},
                     {"name": "delivery", "modules": ["app/lan/delivery"]},
                     {"name": "di", "modules": ["app/lan/di"]}],
          "allowed": {"domain": [], "delivery": ["domain"], "di": ["domain", "delivery"]}}


def _repo2(root, block, files=None):
    repo = Path(root)
    base = {"app/__init__.py": "", "app/lan/__init__.py": "", "app/exp/__init__.py": ""}
    for m in TREE2:
        base[f"{m}/__init__.py"] = ""
    base.update(files or {})
    base["docs/architecture.md"] = _doc_for(block, TREE2)
    for rel, text in base.items():
        (repo / rel).parent.mkdir(parents=True, exist_ok=True)
        (repo / rel).write_text(text, encoding="utf-8")
    (repo / ".importlinter").write_text(imports.render_importlinter(block), encoding="utf-8")
    return repo


def _findings2(repo):
    return gate.check_architecture(repo, "docs/architecture.md", [("app/lan", "*.py"), ("app/exp", "*.py")],
                                   ".importlinter")


class ExtendedRulesRenderTests(unittest.TestCase):
    def test_sin_las_claves_nuevas_la_configuracion_no_cambia(self):
        text = imports.render_importlinter(BLOCK2)
        self.assertNotIn("independent", text)
        self.assertNotIn("ignore_imports", text)
        self.assertNotIn("allow_indirect_imports", text)

    def test_independent_genera_un_contrato_por_par_ordenado_con_importaciones_directas(self):
        block = {**BLOCK2, "independent": [{"modules": ["app/lan", "app/exp"]}]}
        text = imports.render_importlinter(block)
        self.assertIn("[importlinter:contract:independent-app-lan-must-not-import-app-exp]", text)
        self.assertIn("[importlinter:contract:independent-app-exp-must-not-import-app-lan]", text)
        self.assertEqual(text.count("allow_indirect_imports = True"), 2)

    def test_tolerated_se_deriva_a_ignore_imports_del_contrato_de_su_par(self):
        block = {**BLOCK2, "tolerated": [{"from": "delivery", "to": "di", "imports":
                                         ["app/lan/delivery/routes -> app/lan/di/container"]}]}
        text = imports.render_importlinter(block)
        contract = text.split("[importlinter:contract:delivery-must-not-import-di]")[1].split("[importlinter:contract:")[0]
        self.assertIn("ignore_imports =\n    app.lan.delivery.routes -> app.lan.di.container", contract)


class ExtendedRulesGateTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(shutil.which("lint-imports"),
                             "import-linter no está instalado: falla cerrado, no se salta (HR-3)")

    def test_segmentacion_vertical_nombra_fichero_linea_y_regla(self):
        block = {**BLOCK2, "independent": [{"modules": ["app/lan", "app/exp"]}]}
        files = {"app/lan/domain/m.py": "import app.exp.domain\n"}
        with tempfile.TemporaryDirectory() as d:
            found = _findings2(_repo2(d, block, files))
        self.assertEqual(found, ["app/lan/domain/m.py:1: import prohibido por la regla "
                                 "«app.lan must not import app.exp»"])

    def test_segmentacion_sin_cruce_da_verde(self):
        block = {**BLOCK2, "independent": [{"modules": ["app/lan", "app/exp"]}]}
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(_findings2(_repo2(d, block)), [])

    def test_tolerada_presente_pasa_y_una_nueva_no_listada_es_roja(self):
        block = {**BLOCK2, "tolerated": [{"from": "delivery", "to": "di", "imports":
                                         ["app/lan/delivery/routes -> app/lan/di"]}]}
        files = {"app/lan/delivery/routes.py": "import app.lan.di\n"}
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(_findings2(_repo2(d, block, files)), [])
        files["app/lan/delivery/other.py"] = "import app.lan.di\n"
        with tempfile.TemporaryDirectory() as d:
            found = _findings2(_repo2(d, block, files))
        self.assertEqual(found, ["app/lan/delivery/other.py:1: import prohibido por la regla "
                                 "«delivery must not import di»"])

    def test_tolerada_obsoleta_es_stale(self):
        block = {**BLOCK2, "tolerated": [{"from": "delivery", "to": "di", "imports":
                                         ["app/lan/delivery/routes -> app/lan/di"]}]}
        with tempfile.TemporaryDirectory() as d:
            found = _findings2(_repo2(d, block))
        self.assertEqual(len(found), 1)
        self.assertIn("STALE en tolerated", found[0])
        self.assertIn("app.lan.delivery.routes -> app.lan.di", found[0])

    def test_pureza_nombra_fichero_linea_y_framework(self):
        block = {**BLOCK2, "pure": {"layers": ["domain"], "forbidden_externals": ["fastapi", "sqlalchemy"]}}
        files = {"app/lan/domain/m.py": "x = 1\nfrom fastapi import APIRouter\n",
                 "app/exp/domain/n.py": "import sqlalchemy.orm\n",
                 "app/lan/delivery/ok.py": "import fastapi\n"}
        with tempfile.TemporaryDirectory() as d:
            found = sorted(_findings2(_repo2(d, block, files)))
        self.assertEqual(found, ["app/exp/domain/n.py:1: la capa pura «domain» importa el framework «sqlalchemy»",
                                 "app/lan/domain/m.py:2: la capa pura «domain» importa el framework «fastapi»"])

    def test_pureza_cuenta_los_imports_bajo_type_checking(self):
        block = {**BLOCK2, "pure": {"layers": ["domain"], "forbidden_externals": ["fastapi"]}}
        files = {"app/lan/domain/m.py": "from typing import TYPE_CHECKING\nif TYPE_CHECKING:\n    import fastapi\n"}
        with tempfile.TemporaryDirectory() as d:
            found = _findings2(_repo2(d, block, files))
        self.assertEqual(found, ["app/lan/domain/m.py:3: la capa pura «domain» importa el framework «fastapi»"])

    def test_pureza_de_un_fichero_que_no_se_analiza_falla_cerrado(self):
        block = {**BLOCK2, "pure": {"layers": ["domain"], "forbidden_externals": ["fastapi"]}}
        files = {"app/lan/domain/roto.py": "def (:\n"}
        with tempfile.TemporaryDirectory() as d:
            found = _findings2(_repo2(d, block, files))
        self.assertTrue(any("roto.py" in f and "HR-3" in f for f in found), found)


class LinterMissingTests(unittest.TestCase):
    def test_linter_no_instalado_falla_cerrado(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d)
            found = imports.check_imports(repo, BLOCK, ".importlinter", command="no-such-linter-xyz")
            self.assertTrue(any("no está instalado" in f and "HR-3" in f for f in found), found)

    def test_linter_que_falla_sin_salida_legible_no_pasa_en_silencio(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d)
            stub = Path(d) / "stub-linter"
            stub.write_text("#!/bin/sh\necho boom >&2\nexit 3\n", encoding="utf-8")
            stub.chmod(0o755)
            found = imports.check_imports(repo, BLOCK, ".importlinter", command=str(stub))
            self.assertTrue(any("código 3" in f and "boom" in f for f in found), found)


if __name__ == "__main__":
    os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")
    unittest.main()
