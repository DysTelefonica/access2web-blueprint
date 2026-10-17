#!/usr/bin/env python3
# test_architecture_imports_ts.py — reglas de imports en TypeScript (HR-68, #714)
"""Config derivada del bloque, divergencia, linter ausente, import prohibido e
import roto con dependency-cruiser REAL (se instala en CI; si falta, el test
falla, no se salta)."""

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


ts = _load("architecture_imports_ts")
gate = _load("check_docs_architecture")

BLOCK = {"layers": [{"name": "web", "modules": ["src/web"]},
                    {"name": "core", "modules": ["src/core"]}],
         "allowed": {"web": ["core"], "core": []}}
DOC = ("# Arquitectura\n\n## Árbol de módulos\n\n| Módulo | Responsabilidad |\n|---|---|\n"
       "| `src/web` | UI |\n| `src/core` | reglas |\n\n## Capas y dependencias\n\n"
       "```architecture\n" + json.dumps(BLOCK) + "\n```\n")
ROOTS = [("src", "*.ts")]


def _repo(root, core_source="export const m = 1;\n", config=True):
    repo = Path(root)
    files = {"src/web/view.ts": 'import { m } from "../core/model";\nexport const v = m;\n',
             "src/core/model.ts": core_source, "docs/architecture.md": DOC}
    for rel, text in files.items():
        (repo / rel).parent.mkdir(parents=True, exist_ok=True)
        (repo / rel).write_text(text, encoding="utf-8")
    if config:
        (repo / ".dependency-cruiser.cjs").write_text(ts.render_depcruise(BLOCK), encoding="utf-8")
    return repo


def _findings(repo):
    return gate.check_architecture(repo, "docs/architecture.md", ROOTS,
                                   ".dependency-cruiser.cjs", "typescript")


class RenderTests(unittest.TestCase):
    def test_solo_los_pares_no_permitidos_son_reglas_forbidden(self):
        text = ts.render_depcruise(BLOCK)
        self.assertIn('"name": "core-must-not-import-web"', text)
        self.assertIn('"path": "^src/core/"', text)
        self.assertNotIn("web-must-not-import-core", text)
        self.assertIn("no-unresolvable-local-imports", text)

    def test_un_fichero_casa_exacto_y_varios_modulos_se_alternan(self):
        block = {"layers": [{"name": "a", "modules": ["src/main.ts", "src/a"]},
                            {"name": "b", "modules": ["src/b"]}], "allowed": {"a": [], "b": []}}
        text = ts.render_depcruise(block)
        self.assertIn("^(src/main\\\\.ts$|src/a/)", text)

    def test_la_salida_es_determinista(self):
        self.assertEqual(ts.render_depcruise(BLOCK), ts.render_depcruise(BLOCK))


class GateTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(shutil.which("depcruise"),
                             "dependency-cruiser no está instalado: falla cerrado, no se salta (HR-3)")

    def test_codigo_conforme_da_verde(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(_findings(_repo(d)), [])

    def test_import_prohibido_nombra_fichero_linea_y_regla(self):
        with tempfile.TemporaryDirectory() as d:
            found = _findings(_repo(d, core_source='\nimport { v } from "../web/view";\nexport const m = v;\n'))
        self.assertEqual(found, ["src/core/model.ts:2: import prohibido por la regla "
                                 "«core-must-not-import-web»"])

    def test_import_local_roto_no_pasa_en_verde(self):
        with tempfile.TemporaryDirectory() as d:
            found = _findings(_repo(d, core_source='import { q } from "./nope";\nexport const m = q;\n'))
        self.assertEqual(found, ["src/core/model.ts:1: import prohibido por la regla "
                                 "«no-unresolvable-local-imports»"])

    def test_regla_del_documento_sin_reflejo_en_el_linter_es_roja(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d)
            block = {**BLOCK, "allowed": {"web": ["core"], "core": ["web"]}}
            (repo / "docs" / "architecture.md").write_text(
                DOC.replace(json.dumps(BLOCK), json.dumps(block)), encoding="utf-8")
            found = _findings(repo)
            self.assertTrue(any(".dependency-cruiser.cjs" in f and "no coincide" in f for f in found), found)

    def test_write_config_regenera_la_configuracion_desde_el_documento(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, config=False)
            problems = gate.write_config(repo, "docs/architecture.md", ".dependency-cruiser.cjs", "typescript")
            self.assertEqual(problems, [])
            self.assertEqual(_findings(repo), [])


class LinterMissingTests(unittest.TestCase):
    def test_linter_no_instalado_falla_cerrado(self):
        with tempfile.TemporaryDirectory() as d:
            found = ts.check_imports(_repo(d), BLOCK, ".dependency-cruiser.cjs", command="no-such-linter-xyz")
            self.assertTrue(any("no está instalado" in f and "HR-3" in f for f in found), found)

    def test_linter_que_no_analiza_ningun_modulo_no_pasa_en_silencio(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d)
            stub = Path(d) / "stub-linter"
            stub.write_text('#!/bin/sh\necho \'{"summary": {"totalCruised": 0, "violations": []}}\'\n',
                            encoding="utf-8")
            stub.chmod(0o755)
            found = ts.check_imports(repo, BLOCK, ".dependency-cruiser.cjs", command=str(stub))
            self.assertTrue(any("no analizó ningún módulo" in f for f in found), found)

    def test_salida_ilegible_del_linter_es_roja(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d)
            stub = Path(d) / "stub-linter"
            stub.write_text("#!/bin/sh\necho boom\n", encoding="utf-8")
            stub.chmod(0o755)
            found = ts.check_imports(repo, BLOCK, ".dependency-cruiser.cjs", command=str(stub))
            self.assertTrue(any("informe legible" in f for f in found), found)

    def test_sin_ficheros_fuente_falla_cerrado(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d)
            for f in list((repo / "src").rglob("*.ts")):
                f.unlink()
            found = ts.check_imports(repo, BLOCK, ".dependency-cruiser.cjs", command="sh")
            self.assertTrue(any("<empty" in f for f in found), found)


if __name__ == "__main__":
    unittest.main()
