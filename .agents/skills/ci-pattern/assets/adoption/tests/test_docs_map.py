"""HR-63 (#709): el mapa del código hub+spokes se exige por declaración
`docs_map` en el contrato de adopción; sin declaración avisa y sale 0."""

import contextlib
import importlib.machinery
import importlib.util
import io
import json
import tempfile
import unittest
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[3]
CLI_PATH = SKILL_ROOT / "assets" / "bin" / "ci-pattern"


def _load_cli():
    spec = importlib.util.spec_from_loader(
        "ci_pattern_cli_docs_map",
        importlib.machinery.SourceFileLoader("ci_pattern_cli_docs_map", str(CLI_PATH)),
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


CLI = _load_cli()
DECL = {"hub": "docs/CODEBASE-GUIDE.md", "docs_path": "docs/codebase",
        "code_root": "app", "code_glob": "*.py", "job": "docs-codeguide"}
HUB = "# Guía\n\n| Página | Trabajo |\n|---|---|\n| [Mapa](codebase/map.md) | dónde |\n"
WORKFLOW = "jobs:\n  docs-codeguide:\n    runs-on: ubuntu-latest\n"


def _write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


POLICY = {"required_jobs": ["unit", "docs-codeguide"]}


def _repo(root, decl_overrides=None, hub=HUB, page="`app/users.py` usuarios\n",
          workflow=WORKFLOW, policy=POLICY):
    repo = Path(root)
    _write(repo / "app" / "users.py", "x = 1\n")
    if hub is not None:
        _write(repo / "docs" / "CODEBASE-GUIDE.md", hub)
    if page is not None:
        _write(repo / "docs" / "codebase" / "map.md", page)
    if workflow is not None:
        _write(repo / ".github" / "workflows" / "ci.yml", workflow)
    if policy is not None:
        _write(repo / ".github" / "required-jobs-policy.json", json.dumps(policy))
    decl = {**DECL, **(decl_overrides or {})}
    return repo, {"docs_map": decl}


class DocsMapTests(unittest.TestCase):
    def _findings(self, **kw):
        with tempfile.TemporaryDirectory() as d:
            repo, contract = _repo(d, **kw)
            return CLI._check_docs_map(repo, contract)

    def test_conforme_da_verde(self):
        self.assertEqual(self._findings(), [])

    def test_sin_declaracion_avisa_por_stderr_y_sale_verde(self):
        with tempfile.TemporaryDirectory() as d:
            buf = io.StringIO()
            with contextlib.redirect_stderr(buf):
                self.assertEqual(CLI._check_docs_map(Path(d), {}), [])
            self.assertIn("HR-63: docs_map no declarado", buf.getvalue())

    def test_declaracion_incompleta_falla_nombrando_la_clave(self):
        with tempfile.TemporaryDirectory() as d:
            repo, _ = _repo(d)
            found = CLI._check_docs_map(repo, {"docs_map": {"hub": "x.md"}})
            self.assertTrue(any("code_root" in f for f in found), found)

    def test_hub_ausente_falla_cerrado(self):
        found = self._findings(hub=None)
        self.assertTrue(any("CODEBASE-GUIDE.md" in f for f in found), found)

    def test_enlace_del_hub_a_spoke_inexistente(self):
        found = self._findings(hub=HUB + "[Roto](codebase/ghost.md)\n")
        self.assertTrue(any("codebase/ghost.md" in f for f in found), found)

    def test_spoke_no_enlazado_desde_el_hub(self):
        found = self._findings(hub="# Guía\n\nsin enlaces\n")
        self.assertTrue(any("map.md" in f and "hub" in f for f in found), found)

    def test_modulo_sin_documentar_falla(self):
        found = self._findings(page="sin modulos\n")
        self.assertTrue(any("users" in f for f in found), found)

    def test_ruta_citada_inexistente_falla(self):
        found = self._findings(page="`app/users.py` y `app/ghost.py`\n")
        self.assertTrue(any("app/ghost.py" in f for f in found), found)

    def test_job_requerido_ausente_del_workflow(self):
        found = self._findings(workflow="jobs:\n  lint:\n    runs-on: x\n")
        self.assertTrue(any("docs-codeguide" in f and "job" in f for f in found), found)

    def test_job_fuera_de_required_jobs_falla(self):
        found = self._findings(policy={"required_jobs": ["unit"]})
        self.assertTrue(any("required_jobs" in f and "docs-codeguide" in f for f in found), found)

    def test_sin_politica_de_required_jobs_falla_cerrado(self):
        found = self._findings(policy=None)
        self.assertTrue(any("required-jobs-policy.json" in f for f in found), found)

    def test_ruta_declarada_ausente_que_aparece_falla_hasta_actualizar_la_pagina(self):
        page = "`app/users.py`\n\n## No existe en este repo\n\n- dashboard/\n"
        with tempfile.TemporaryDirectory() as d:
            repo, contract = _repo(d, page=page)
            self.assertEqual(CLI._check_docs_map(repo, contract), [])
            (repo / "dashboard").mkdir()
            found = CLI._check_docs_map(repo, contract)
            self.assertTrue(any("dashboard/" in f and "ausente" in f for f in found), found)


ROOTS = [{"root": "app", "glob": "*.py"}, {"root": "templates", "glob": "*.html"}]


def _roots_decl(roots=None):
    decl = {k: v for k, v in DECL.items() if k not in ("code_root", "code_glob")}
    decl["code_roots"] = ROOTS if roots is None else roots
    return decl


class DocsMapRootsTests(unittest.TestCase):
    """#727: docs_map admite varias raíces de código (lista de raíz + glob)."""

    def _findings(self, decl, page="`app/users.py` y la plantilla home\n", templates=True):
        with tempfile.TemporaryDirectory() as d:
            repo, _ = _repo(d, page=page)
            if templates:
                _write(repo / "templates" / "home.html", "<p>hola</p>\n")
            return CLI._check_docs_map(repo, {"docs_map": decl})

    def test_dos_raices_conformes_dan_verde(self):
        self.assertEqual(self._findings(_roots_decl()), [])

    def test_modulo_sin_describir_en_la_segunda_raiz_nombra_raiz_y_modulo(self):
        found = self._findings(_roots_decl(), page="`app/users.py`\n")
        self.assertTrue(any("home" in f and "templates" in f for f in found), found)

    def test_modulo_sin_describir_en_la_primera_raiz_nombra_raiz_y_modulo(self):
        found = self._findings(_roots_decl(), page="la plantilla home\n")
        self.assertTrue(any("users" in f and "app" in f for f in found), found)

    def test_la_ruta_citada_inexistente_se_informa_una_sola_vez(self):
        page = "`app/users.py`, home y `app/ghost.py`\n"
        found = [f for f in self._findings(_roots_decl(), page=page) if "app/ghost.py" in f]
        self.assertEqual(len(found), 1, found)

    def test_raiz_inexistente_falla_cerrado(self):
        decl = _roots_decl(ROOTS + [{"root": "static", "glob": "*.js"}])
        found = self._findings(decl)
        self.assertTrue(any("static" in f and "empty" in f for f in found), found)

    def test_declaracion_escalar_existente_sigue_en_verde(self):
        self.assertEqual(self._findings(dict(DECL), page="`app/users.py`\n", templates=False), [])

    def test_lista_vacia_falla_nombrando_la_clave(self):
        found = self._findings(_roots_decl([]))
        self.assertTrue(any("code_roots" in f for f in found), found)

    def test_elemento_sin_root_o_sin_glob_falla_nombrando_la_clave(self):
        for bad in ({"glob": "*.py"}, {"root": "app"}, {"root": "app", "glob": ""}, "app"):
            found = self._findings(_roots_decl([bad]))
            self.assertTrue(any("code_roots" in f for f in found), (bad, found))

    def test_mezcla_de_escalar_y_lista_falla_nombrando_las_claves(self):
        decl = {**_roots_decl(), "code_root": "app", "code_glob": "*.py"}
        found = self._findings(decl)
        self.assertTrue(any("code_roots" in f and "code_root" in f for f in found), found)


if __name__ == "__main__":
    unittest.main()
