"""HR-82 (#1382): las tablas de checks escritas a mano, contra los jobs reales.

La fase 5 mide la tabla del documento operativo GENERADO (`docs/ci-pattern-flow.md`);
las tablas escritas a mano no se medían. El caso vivo es `DysTelefonica/cadete`: la
tabla de `docs/codebase/testing-and-ci.md` atribuye `php-lint` a `ci.yml` (el job
vive en `php-lint-push.yml`), omite `ci-images` (que sí es un job requerido) y una
auditoría le añadió un job que no define ningún workflow. Con la capa declarada,
los tres son rojos; la tabla alineada, verde.
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

CI_YML = """name: ci
on:
  pull_request:
jobs:
  phpunit:
    name: PHPUnit Integration Tests
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - run: echo phpunit
  playwright-e2e:
    name: Playwright E2E Tests
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - run: echo playwright
  governance-seals:
    name: governance-seals
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - run: echo seals
  docs-gates:
    name: docs-gates
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - run: echo gates
  docs-commands:
    name: docs-commands (${{ github.event_name }})
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - run: echo commands
  pr-gate:
    name: pr-gate
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - run: echo gate
  ci-images:
    name: ci-images
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - run: echo images
  required:
    name: required
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - run: echo required
"""

PHP_LINT_YML = """name: php-lint-push
on:
  push:
    branches: [main]
jobs:
  php-lint:
    name: PHP Lint (php -l)
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - run: echo lint
"""

POLICY = {
    "required_jobs": [
        "phpunit",
        "playwright-e2e",
        "governance-seals",
        "docs-gates",
        "docs-commands",
        "pr-gate",
        "ci-images",
    ],
}

CABECERA = """## Los jobs de CI

`.github/workflows/ci.yml` define los jobs del CI.

| Job | Nombre del check | Qué hace |
|---|---|---|
"""

FILAS = """| `phpunit` | `PHPUnit Integration Tests` | Levanta MySQL en Docker y ejecuta PHPUnit. |
| `playwright-e2e` | `Playwright E2E Tests` | Levanta la aplicación y ejecuta el navegador. |
| `php-lint` | `PHP Lint (php -l)` | Ejecuta el preflight del lint. |
| `pr-gate` | `pr-gate` | Comprueba la issue enlazada. |
| `docs-gates` | `docs-gates` | Fase 5 de la adopción. |
| `governance-seals` | `governance-seals` | Fase 3 de la adopción. |
| `docs-commands` | `docs-commands (pull_request)` | Ejecuta los comandos documentados. |
| `required` | `required` | Agregador de los jobs anteriores. |
"""

# El job inventado de la auditoría: no lo define ningún workflow.
FILA_INVENTADA = "| `security-scan` | `Security Scan` | Escanea las dependencias. |\n"

# La tabla alineada: `php-lint` sale (vive en otro workflow) y entra `ci-images`.
FILAS_ALINEADAS = """| `phpunit` | `PHPUnit Integration Tests` | Levanta MySQL en Docker y ejecuta PHPUnit. |
| `playwright-e2e` | `Playwright E2E Tests` | Levanta la aplicación y ejecuta el navegador. |
| `pr-gate` | `pr-gate` | Comprueba la issue enlazada. |
| `docs-gates` | `docs-gates` | Fase 5 de la adopción. |
| `governance-seals` | `governance-seals` | Fase 3 de la adopción. |
| `docs-commands` | `docs-commands (pull_request)` | Ejecuta los comandos documentados. |
| `ci-images` | `ci-images` | Fija las imágenes del CI por digest. |
| `required` | `required` | Agregador de los jobs anteriores. |
"""


def _load_cli():
    spec = importlib.util.spec_from_loader(
        "ci_pattern_cli_checks_tables",
        importlib.machinery.SourceFileLoader("ci_pattern_cli_checks_tables", str(CLI_PATH)),
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


CLI = _load_cli()


def _write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _repo(root, filas=FILAS, seccion="Los jobs de CI", workflow="ci.yml"):
    """Un repo con la forma del de cadete: la tabla, los workflows y la política."""
    repo = Path(root)
    _write(repo / "docs" / "codebase" / "testing-and-ci.md",
           CABECERA + filas + "\n## Otra sección\n\nNo es una tabla de jobs.\n")
    _write(repo / ".github" / "workflows" / "ci.yml", CI_YML)
    _write(repo / ".github" / "workflows" / "php-lint-push.yml", PHP_LINT_YML)
    _write(repo / ".github" / "required-jobs-policy.json", json.dumps(POLICY, indent=2))
    tabla = {"path": "docs/codebase/testing-and-ci.md", "workflow": workflow}
    if seccion:
        tabla["section"] = seccion
    contrato = {"checks_tables": {"tables": [tabla]}}
    return repo, contrato


class ChecksTablesTests(unittest.TestCase):
    """La tabla declarada se mide contra los jobs reales de los workflows."""

    def _findings(self, **kw):
        with tempfile.TemporaryDirectory() as d:
            repo, contrato = _repo(d, **kw)
            return CLI._check_checks_tables(repo, contrato)

    def test_sin_declaracion_avisa_por_stderr_y_sale_verde(self):
        with tempfile.TemporaryDirectory() as d:
            repo, _ = _repo(d)
            buf = io.StringIO()
            with contextlib.redirect_stderr(buf):
                self.assertEqual(CLI._check_checks_tables(repo, {}), [])
            self.assertIn("HR-82: checks_tables no declarado", buf.getvalue())

    def test_job_inventado_es_rojo(self):
        found = self._findings(filas=FILAS + FILA_INVENTADA)
        self.assertTrue(any("security-scan" in f for f in found), found)

    def test_job_movido_de_workflow_es_rojo(self):
        found = self._findings()
        self.assertTrue(any("php-lint" in f and "php-lint-push.yml" in f for f in found), found)

    def test_job_requerido_omitido_es_rojo(self):
        found = self._findings()
        self.assertTrue(any("ci-images" in f for f in found), found)

    def test_tabla_alineada_da_verde(self):
        self.assertEqual(self._findings(filas=FILAS_ALINEADAS), [])

    def test_tabla_sin_filas_de_job_falla_cerrado(self):
        found = self._findings(filas="| Capa | Dónde | Qué protege |\n|---|---|---|\n| Unitarias | `www/tests` | Contratos. |\n")
        self.assertTrue(any("filas de jobs" in f for f in found), found)

    def test_seccion_que_no_existe_falla(self):
        found = self._findings(seccion="Una sección que no existe")
        self.assertTrue(any("secci" in f for f in found), found)

    def test_workflow_que_no_existe_falla(self):
        found = self._findings(workflow="no-existe.yml")
        self.assertTrue(any("no-existe.yml" in f for f in found), found)

    def test_declaracion_que_no_es_objeto_falla(self):
        with tempfile.TemporaryDirectory() as d:
            repo, _ = _repo(d)
            found = CLI._check_checks_tables(repo, {"checks_tables": "tablas"})
        self.assertTrue(any("objeto" in f for f in found), found)

    def test_el_esquema_admite_la_clave_opcional(self):
        schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        self.assertIn("checks_tables", schema["properties"])
        self.assertNotIn("checks_tables", schema.get("required", []))
        self.assertIn("tables", schema["properties"]["checks_tables"]["properties"])


if __name__ == "__main__":
    unittest.main()
