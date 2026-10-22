#!/usr/bin/env python3
# test_check_docs_architecture.py — gate docs-architecture (HR-68, #714)
"""Árbol↔código en los dos sentidos, capas y fail-closed. Solo stdlib."""

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_docs_architecture.py"
_SPEC = importlib.util.spec_from_file_location("check_docs_architecture", GATE)
gate = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(gate)

BLOCK = {"layers": [{"name": "web", "modules": ["app/routes"]},
                    {"name": "core", "modules": ["app/domain"]}],
         "allowed": {"web": ["core"], "core": []}}
ROOTS = [("app", "*.py")]


def _doc(tree=("app/routes", "app/domain"), block=BLOCK):
    rows = "".join(f"| `{p}` | hace algo |\n" for p in tree)
    body = f"```architecture\n{json.dumps(block)}\n```\n" if block is not None else ""
    return ("# Arquitectura\n\n## Árbol de módulos\n\n| Módulo | Responsabilidad |\n|---|---|\n"
            f"{rows}\n## Capas y dependencias\n\n{body}")


def _repo(root, doc, modules=("routes", "domain")):
    repo = Path(root)
    for name in modules:
        (repo / "app" / name).mkdir(parents=True, exist_ok=True)
        (repo / "app" / name / "x.py").write_text("x = 1\n", encoding="utf-8")
    if doc is not None:
        (repo / "docs").mkdir(exist_ok=True)
        (repo / "docs" / "architecture.md").write_text(doc, encoding="utf-8")
    return repo


class ExtendedRulesValidationTests(unittest.TestCase):
    """Claves opcionales del bloque (#749): forma válida y rechazos."""

    def _findings(self, **extra):
        block = {**BLOCK, **extra}
        with tempfile.TemporaryDirectory() as d:
            return gate.check_architecture(_repo(d, _doc(block=block)), "docs/architecture.md", ROOTS)

    def test_las_tres_claves_bien_formadas_dan_verde(self):
        self.assertEqual(self._findings(
            independent=[{"modules": ["app/a", "app/b"]}],
            pure={"layers": ["core"], "forbidden_externals": ["fastapi"]},
            tolerated=[{"from": "core", "to": "web", "imports": ["app/domain/x -> app/routes/y"]}]), [])

    def test_independent_con_un_solo_modulo_es_rojo(self):
        found = self._findings(independent=[{"modules": ["app/a"]}])
        self.assertTrue(any("«independent»" in f for f in found), found)

    def test_pure_con_capa_no_declarada_es_roja(self):
        found = self._findings(pure={"layers": ["nube"], "forbidden_externals": ["fastapi"]})
        self.assertTrue(any("«pure» nombra la capa nube" in f for f in found), found)

    def test_pure_sin_externos_es_roja(self):
        found = self._findings(pure={"layers": ["core"], "forbidden_externals": []})
        self.assertTrue(any("«pure» necesita" in f for f in found), found)

    def test_tolerated_de_una_dependencia_ya_permitida_es_roja(self):
        found = self._findings(tolerated=[{"from": "web", "to": "core", "imports": ["a -> b"]}])
        self.assertTrue(any("no tiene sentido" in f for f in found), found)

    def test_tolerated_con_capa_no_declarada_o_forma_invalida_es_roja(self):
        found = self._findings(tolerated=[{"from": "core", "to": "nube", "imports": ["a -> b"]},
                                          {"from": "core", "to": "web", "imports": ["sin flecha"]}])
        self.assertTrue(any("capa no declarada" in f for f in found), found)
        self.assertTrue(any("origen -> destino" in f for f in found), found)

    def test_claves_que_no_son_listas_son_rojas(self):
        found = self._findings(independent={"modules": ["a", "b"]}, tolerated="x")
        self.assertTrue(any("«independent» debe ser una lista" in f for f in found), found)
        self.assertTrue(any("«tolerated» debe ser una lista" in f for f in found), found)


class UnlayeredFilesTests(unittest.TestCase):
    """``covered_roots`` y ``exempt`` (#771): todo fichero fuente bajo un paquete cubierto
    cae en una capa o en ``exempt``; sin la clave nada cambia. Disposición como la de un
    monolito modular: los code-roots son las aplicaciones y los módulos, sus carpetas de capa."""

    BLOCK = {"layers": [{"name": "web", "modules": ["app/mods/a/web"]},
                        {"name": "core", "modules": ["app/mods/a/core"]}],
             "allowed": {"web": ["core"], "core": []}}
    TREE = ("app/mods/a/web", "app/mods/a/core")
    CODE_ROOTS = [("app/mods/a", "*.py")]

    def _findings(self, files=None, **extra):
        block = {**self.BLOCK, **extra}
        with tempfile.TemporaryDirectory() as d:
            repo = Path(d)
            for rel in ("app/mods/a/web/x.py", "app/mods/a/core/x.py"):
                (repo / rel).parent.mkdir(parents=True, exist_ok=True)
                (repo / rel).write_text("x = 1\n", encoding="utf-8")
            (repo / "docs").mkdir()
            (repo / "docs" / "architecture.md").write_text(_doc(self.TREE, block), encoding="utf-8")
            for rel, text in (files or {}).items():
                (repo / rel).parent.mkdir(parents=True, exist_ok=True)
                (repo / rel).write_text(text, encoding="utf-8")
            return gate.check_architecture(repo, "docs/architecture.md", self.CODE_ROOTS)

    def test_un_fichero_suelto_bajo_el_paquete_cubierto_es_rojo_y_se_nombra(self):
        found = self._findings({"app/loose_helper.py": "x = 1\n"}, covered_roots=["app"])
        self.assertEqual(found, ["el fichero app/loose_helper.py no pertenece a ninguna capa ni a «exempt»"])

    def test_sin_la_clave_el_mismo_fichero_pasa(self):
        self.assertEqual(self._findings({"app/loose_helper.py": "x = 1\n"}), [])

    def test_los_ficheros_de_los_modulos_de_una_capa_pertenecen_a_ella(self):
        files = {"app/mods/a/web/sub/deep.py": "x = 1\n", "app/mods/a/core/more.py": "x = 1\n"}
        self.assertEqual(self._findings(files, covered_roots=["app"]), [])

    def test_un_fichero_exento_pasa(self):
        found = self._findings({"app/main.py": "x = 1\n"}, covered_roots=["app"], exempt=["app/main.py"])
        self.assertEqual(found, [])

    def test_exempt_admite_patrones(self):
        files = {"app/scripts/a.py": "x = 1\n", "app/scripts/b.py": "x = 1\n"}
        self.assertEqual(self._findings(files, covered_roots=["app"], exempt=["app/scripts/*.py"]), [])

    def test_una_entrada_de_exempt_que_no_casa_nada_es_stale(self):
        found = self._findings(covered_roots=["app"], exempt=["app/ya_no_existe.py"])
        self.assertEqual(found, ["«exempt» declara el patrón app/ya_no_existe.py y no casa ningún fichero"])

    def test_el_marcador_de_paquete_por_encima_de_los_modulos_no_cuenta(self):
        files = {"app/__init__.py": "", "app/mods/__init__.py": "", "app/mods/a/__init__.py": "",
                 "app/mods/a/web/__init__.py": ""}
        self.assertEqual(self._findings(files, covered_roots=["app"]), [])

    def test_un_init_dentro_de_un_directorio_sin_capa_es_rojo(self):
        found = self._findings({"app/otro/__init__.py": ""}, covered_roots=["app"])
        self.assertEqual(found, ["el fichero app/otro/__init__.py no pertenece a ninguna capa ni a «exempt»"])

    def test_forma_invalida_de_las_claves_es_roja(self):
        found = self._findings(covered_roots="app", exempt="app/main.py")
        self.assertTrue(any("«covered_roots» debe ser una lista" in f for f in found), found)
        self.assertTrue(any("«exempt» debe ser una lista" in f for f in found), found)

    def test_un_paquete_cubierto_que_no_existe_es_rojo(self):
        found = self._findings(covered_roots=["nada"])
        self.assertEqual(found, ["«covered_roots» declara nada, que no es un directorio"])


class TreeAndLayersTests(unittest.TestCase):
    def _findings(self, doc, modules=("routes", "domain"), roots=ROOTS):
        with tempfile.TemporaryDirectory() as d:
            return gate.check_architecture(_repo(d, doc, modules), "docs/architecture.md", roots)

    def test_arbol_que_cuadra_da_verde(self):
        self.assertEqual(self._findings(_doc()), [])

    def test_modulo_descrito_inexistente_es_rojo_nombrandolo(self):
        found = self._findings(_doc(tree=("app/routes", "app/domain", "app/ghost")))
        self.assertTrue(any("app/ghost" in f and "árbol → código" in f for f in found), found)

    def test_modulo_existente_sin_describir_es_rojo_nombrandolo(self):
        found = self._findings(_doc(), modules=("routes", "domain", "billing"))
        self.assertTrue(any("app/billing" in f and "código → árbol" in f for f in found), found)

    def test_arbol_vacio_falla_cerrado(self):
        found = self._findings(_doc(tree=()))
        self.assertTrue(any("<empty árbol" in f for f in found), found)

    def test_code_root_vacio_falla_cerrado(self):
        found = self._findings(_doc(), modules=())
        self.assertTrue(any("<empty app/>" in f for f in found), found)

    def test_init_py_no_es_un_modulo(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, _doc())
            (repo / "app" / "__init__.py").write_text("", encoding="utf-8")
            self.assertEqual(gate.check_architecture(repo, "docs/architecture.md", ROOTS), [])

    def test_las_subsecciones_del_arbol_cuentan_hasta_un_titulo_de_su_nivel(self):
        """Regresión #751: el árbol bajo subtítulos `###` se lee entero."""
        doc = ("# Arquitectura\n\n## Árbol de módulos\n\n### Raíz\n\n| Módulo | Responsabilidad |\n|---|---|\n"
               "| `app/routes` | HTTP |\n\n### Núcleo\n\n| Módulo | Responsabilidad |\n|---|---|\n"
               "| `app/domain` | reglas |\n\n## Otra sección\n\n| `app/ghost` | no es del árbol |\n\n"
               "## Capas y dependencias\n\n```architecture\n" + json.dumps(BLOCK) + "\n```\n")
        self.assertEqual(gate.parse_tree(doc), ["app/routes", "app/domain"])
        self.assertEqual(self._findings(doc), [])

    def test_documento_ausente_falla_nombrandolo(self):
        found = self._findings(None)
        self.assertEqual(found, ["falta el documento docs/architecture.md"])

    def test_sin_bloque_de_capas_es_rojo(self):
        found = self._findings(_doc(block=None))
        self.assertTrue(any("bloque cercado «architecture»" in f for f in found), found)

    def test_bloque_no_json_es_rojo(self):
        doc = _doc(block=None) + "```architecture\n{no json\n```\n"
        self.assertTrue(any("JSON" in f for f in self._findings(doc)))

    def test_modulo_nuevo_fuera_de_toda_capa_es_rojo(self):
        doc = _doc(tree=("app/routes", "app/domain", "app/billing"))
        found = self._findings(doc, modules=("routes", "domain", "billing"))
        self.assertTrue(any("app/billing" in f and "ninguna capa" in f for f in found), found)

    def test_modulo_en_dos_capas_es_rojo(self):
        block = {**BLOCK, "layers": BLOCK["layers"] + [{"name": "dup", "modules": ["app/domain"]}],
                 "allowed": {**BLOCK["allowed"], "dup": []}}
        found = self._findings(_doc(block=block))
        self.assertTrue(any("app/domain" in f and "varias capas" in f for f in found), found)

    def test_capa_que_declara_un_modulo_fuera_del_arbol_es_rojo(self):
        block = {**BLOCK, "layers": [{"name": "web", "modules": ["app/routes", "app/ghost"]},
                                     BLOCK["layers"][1]]}
        found = self._findings(_doc(block=block))
        self.assertTrue(any("web" in f and "app/ghost" in f for f in found), found)

    def test_capa_sin_regla_es_rojo(self):
        block = {**BLOCK, "allowed": {"web": ["core"]}}
        self.assertTrue(any("core" in f and "no tiene regla" in f
                            for f in self._findings(_doc(block=block))))

    def test_regla_que_nombra_una_capa_no_declarada_es_roja(self):
        block = {**BLOCK, "allowed": {"web": ["core", "ghost"], "core": []}}
        self.assertTrue(any("ghost" in f for f in self._findings(_doc(block=block))))

    def test_el_arbol_dentro_de_un_bloque_cercado_no_cuenta(self):
        doc = "# A\n\n## Árbol de módulos\n\n```text\n| `app/ghost` | x |\n```\n"
        self.assertTrue(any("<empty árbol" in f for f in self._findings(doc)))

    def test_varias_raices_usan_rutas_completas(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, _doc(tree=("app/routes", "app/domain", "web/page.html")))
            (repo / "web").mkdir()
            (repo / "web" / "page.html").write_text("<p>x</p>", encoding="utf-8")
            block = {"layers": [{"name": "web", "modules": ["app/routes", "web/page.html"]},
                                BLOCK["layers"][1]], "allowed": BLOCK["allowed"]}
            (repo / "docs" / "architecture.md").write_text(
                _doc(tree=("app/routes", "app/domain", "web/page.html"), block=block), encoding="utf-8")
            roots = [("app", "*.py"), ("web", "*.html")]
            self.assertEqual(gate.check_architecture(repo, "docs/architecture.md", roots), [])


class CliTests(unittest.TestCase):
    def _run(self, repo, *extra):
        return subprocess.run([sys.executable, str(GATE), str(repo), *extra], capture_output=True,
                              text=True, check=False, encoding="utf-8")

    def test_exit_0_en_verde_y_1_en_rojo_y_2_con_entrada_invalida(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, _doc())
            self.assertEqual(self._run(repo, "--code-root", "app", "--code-glob", "*.py").returncode, 0)
            (repo / "app" / "billing").mkdir()
            (repo / "app" / "billing" / "x.py").write_text("x = 1\n", encoding="utf-8")
            red = self._run(repo, "--code-root", "app", "--code-glob", "*.py")
            self.assertEqual(red.returncode, 1)
            self.assertIn("app/billing", red.stderr)
            self.assertEqual(self._run(repo).returncode, 2)
            self.assertEqual(self._run(repo, "--code-root", "app").returncode, 2)
            self.assertEqual(self._run("/no/existe", "--code-root", "a", "--code-glob", "*.py").returncode, 2)


if __name__ == "__main__":
    unittest.main()
