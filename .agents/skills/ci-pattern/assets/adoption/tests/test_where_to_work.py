"""HR-77 (#1201): el modo de la página «dónde trabajo» lo declara el contrato.

La clave `docs_where_to_work` es **opcional**: sin ella el gate no se mide y
nadie cambia de comportamiento. Con ella, `mode` dice si la página se exige
(`exigido`) o solo se avisa (`aviso`). El modo es dato del contrato, no
cableado del consumer.
"""

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
SCHEMA_PATH = SKILL_ROOT / "assets" / "adoption" / "adoption-contract.schema.json"

PAGE = """# Dónde trabajar

| Tarea | Lee | Toca | Prueba |
|---|---|---|---|
| `form` | `docs/codebase/modulo.md` | `app/modulo` | `tests/test_modulo.py` |
| `ci` | `docs/codebase/ci-cd.md` | `.github/workflows/ci.yml` | `make verify` |
| `schema` | `docs/codebase/esquema.md` | `app/schema.sql` | `tests/test_schema.py` |
| `local` | `docs/CODEBASE-GUIDE.md` | `skills/local/SKILL.md` | `make lint` |
"""


def _load_cli():
    spec = importlib.util.spec_from_loader(
        "ci_pattern_cli_where_to_work",
        importlib.machinery.SourceFileLoader("ci_pattern_cli_where_to_work", str(CLI_PATH)),
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


CLI = _load_cli()


def _write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _repo(root, page=PAGE, mode="exigido"):
    repo = Path(root)
    _write(repo / "docs" / "codebase" / "where-to-work.md", page)
    for name in (
        "docs/codebase/modulo.md",
        "docs/codebase/ci-cd.md",
        "docs/codebase/esquema.md",
        "docs/CODEBASE-GUIDE.md",
    ):
        _write(repo / name, "x\n")
    _write(repo / "app" / "modulo" / "routes.py", "x = 1\n")
    _write(repo / "app" / "schema.sql", "select 1;\n")
    _write(repo / ".github" / "workflows" / "ci.yml", "name: ci\n")
    _write(repo / "skills" / "local" / "SKILL.md", "# skill\n")
    _write(repo / "tests" / "test_modulo.py", "def test_x(): pass\n")
    _write(repo / "tests" / "test_schema.py", "def test_x(): pass\n")
    contract = {} if mode is None else {"docs_where_to_work": {"mode": mode}}
    return repo, contract


class WhereToWorkContractTests(unittest.TestCase):
    """El modo lo declara el contrato; sin declaración, nada cambia."""

    def _findings(self, **kw):
        with tempfile.TemporaryDirectory() as d:
            repo, contract = _repo(d, **kw)
            return CLI._check_where_to_work(repo, contract)

    def test_sin_declaracion_avisa_por_stderr_y_sale_verde(self):
        with tempfile.TemporaryDirectory() as d:
            buf = io.StringIO()
            with contextlib.redirect_stderr(buf):
                self.assertEqual(CLI._check_where_to_work(Path(d), {}), [])
            self.assertIn("HR-77: docs_where_to_work no declarado", buf.getvalue())

    def test_modo_exigido_con_pagina_conforme_da_verde(self):
        self.assertEqual(self._findings(mode="exigido"), [])

    def test_modo_exigido_con_pagina_rota_falla(self):
        found = self._findings(page="# vacía\n", mode="exigido")
        self.assertTrue(any("where-to-work.md" in f for f in found), found)

    def test_modo_aviso_con_pagina_rota_no_falla_pero_avisa(self):
        with tempfile.TemporaryDirectory() as d:
            repo, contract = _repo(d, page="# vacía\n", mode="aviso")
            buf = io.StringIO()
            with contextlib.redirect_stderr(buf):
                self.assertEqual(CLI._check_where_to_work(repo, contract), [])
            self.assertIn("where-to-work.md", buf.getvalue())

    def test_modo_desconocido_falla_cerrado(self):
        found = self._findings(mode="otro")
        self.assertTrue(any("exigido" in f for f in found), found)

    def test_declaracion_que_no_es_objeto_falla(self):
        found = self._findings(mode=None)
        with tempfile.TemporaryDirectory() as d:
            repo, _ = _repo(d)
            found = CLI._check_where_to_work(repo, {"docs_where_to_work": "exigido"})
        self.assertTrue(any("objeto" in f for f in found), found)

    def test_el_esquema_admite_la_clave_opcional(self):
        schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        prop = schema["properties"]["docs_where_to_work"]
        self.assertIn("mode", prop["properties"])
        self.assertNotIn("docs_where_to_work", schema.get("required", []))

    def test_la_clave_ausente_no_altera_el_esquema_del_contrato(self):
        schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        self.assertNotIn("docs_where_to_work", schema.get("required", []))
        self.assertEqual(
            {"exigido", "aviso"},
            set(schema["properties"]["docs_where_to_work"]["properties"]["mode"]["enum"]),
        )


if __name__ == "__main__":
    unittest.main()
