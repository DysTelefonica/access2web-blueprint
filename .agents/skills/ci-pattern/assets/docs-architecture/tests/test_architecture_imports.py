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
