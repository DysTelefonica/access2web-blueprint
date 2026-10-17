"""HR-68 (#714): el documento de arquitectura se exige por declaración
`architecture` en el contrato de adopción; sin declaración avisa y sale 0."""

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
_LOADER = importlib.machinery.SourceFileLoader("ci_pattern_cli_architecture", str(CLI_PATH))
_SPEC = importlib.util.spec_from_loader(_LOADER.name, _LOADER)
CLI = importlib.util.module_from_spec(_SPEC)
_LOADER.exec_module(CLI)

DECL = {"doc": "docs/architecture.md", "code_root": "app", "code_glob": "*.py",
        "job": "docs-architecture", "stack": "python", "linter_config": ".importlinter"}
BLOCK = {"layers": [{"name": "web", "modules": ["app/routes"]},
                    {"name": "core", "modules": ["app/domain"]}],
         "allowed": {"web": ["core"], "core": []}}
DOC = ("# Arquitectura\n\n## Árbol de módulos\n\n| Módulo | Responsabilidad |\n|---|---|\n"
       "| `app/routes` | HTTP |\n| `app/domain` | reglas |\n\n## Capas y dependencias\n\n"
       "```architecture\n" + json.dumps(BLOCK) + "\n```\n")
WORKFLOW = "jobs:\n  docs-architecture:\n    runs-on: ubuntu-latest\n"
POLICY = {"required_jobs": ["unit", "docs-architecture"]}


_IMPORTS_SPEC = importlib.util.spec_from_file_location(
    "architecture_imports", SKILL_ROOT / "assets" / "docs-architecture" / "architecture_imports.py")
IMPORTS = importlib.util.module_from_spec(_IMPORTS_SPEC)
_IMPORTS_SPEC.loader.exec_module(IMPORTS)
_TS_SPEC = importlib.util.spec_from_file_location(
    "architecture_imports_ts", SKILL_ROOT / "assets" / "docs-architecture" / "architecture_imports_ts.py")
TS = importlib.util.module_from_spec(_TS_SPEC)
_TS_SPEC.loader.exec_module(TS)


def _write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _repo(root, doc=DOC, workflow=WORKFLOW, policy=POLICY, modules=("routes", "domain"),
          domain_source="x = 1\n", config=True):
    repo = Path(root)
    _write(repo / "app" / "__init__.py", "")
    for name in modules:
        _write(repo / "app" / name / "__init__.py", "")
        _write(repo / "app" / name / "x.py", domain_source if name == "domain" else "x = 1\n")
    if config:
        _write(repo / ".importlinter", IMPORTS.render_importlinter(BLOCK))
    if doc is not None:
        _write(repo / "docs" / "architecture.md", doc)
    if workflow is not None:
        _write(repo / ".github" / "workflows" / "ci.yml", workflow)
    if policy is not None:
        _write(repo / ".github" / "required-jobs-policy.json", json.dumps(policy))
    return repo


class ArchitectureTests(unittest.TestCase):
    def _findings(self, decl=None, **kw):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, **kw)
            return CLI._check_architecture(repo, {"architecture": {**DECL, **(decl or {})}})

    def test_conforme_da_verde(self):
        self.assertEqual(self._findings(), [])

    def test_sin_declaracion_avisa_por_stderr_y_sale_verde(self):
        buf = io.StringIO()
        with tempfile.TemporaryDirectory() as d, contextlib.redirect_stderr(buf):
            self.assertEqual(CLI._check_architecture(Path(d), {}), [])
        self.assertIn("HR-68: architecture no declarado", buf.getvalue())

    def test_declaracion_incompleta_falla_nombrando_las_claves(self):
        with tempfile.TemporaryDirectory() as d:
            found = CLI._check_architecture(_repo(d), {"architecture": {"doc": "x.md"}})
        self.assertTrue(any("job" in f and "code_root" in f for f in found), found)

    def test_documento_ausente_falla_nombrandolo(self):
        found = self._findings(doc=None)
        self.assertTrue(any("docs/architecture.md" in f for f in found), found)

    def test_modulo_descrito_inexistente_falla(self):
        found = self._findings(doc=DOC.replace("| `app/domain` | reglas |", "| `app/ghost` | x |"))
        self.assertTrue(any("app/ghost" in f and "HR-68" in f for f in found), found)

    def test_modulo_existente_sin_describir_falla(self):
        found = self._findings(modules=("routes", "domain", "billing"))
        self.assertTrue(any("app/billing" in f for f in found), found)

    def test_modulo_nuevo_sin_capa_falla(self):
        doc = DOC.replace("\n## Capas", "| `app/billing` | cobros |\n\n## Capas", 1)
        found = self._findings(doc=doc, modules=("routes", "domain", "billing"))
        self.assertTrue(any("app/billing" in f and "ninguna capa" in f for f in found), found)

    def test_varias_raices_admitidas_por_lista(self):
        decl = {"code_roots": [{"root": "app", "glob": "*.py"}]}
        with tempfile.TemporaryDirectory() as d:
            contract = {"architecture": {k: DECL[k] for k in ("doc", "job", "stack", "linter_config")} | decl}
            self.assertEqual(CLI._check_architecture(_repo(d), contract), [])

    def test_mezcla_de_escalar_y_lista_falla_nombrando_hr_68(self):
        found = self._findings({"code_roots": [{"root": "app", "glob": "*.py"}]})
        self.assertTrue(any("HR-68" in f and "architecture" in f and "code_roots" in f for f in found), found)

    def test_declaracion_sin_stack_ni_linter_falla_nombrando_las_claves(self):
        with tempfile.TemporaryDirectory() as d:
            decl = {k: v for k, v in DECL.items() if k not in ("stack", "linter_config")}
            found = CLI._check_architecture(_repo(d), {"architecture": decl})
        self.assertTrue(any("stack" in f and "linter_config" in f for f in found), found)

    def test_stack_no_soportado_falla_cerrado(self):
        found = self._findings({"stack": "cobol"})
        self.assertTrue(any("stack" in f and "no soportado" in f for f in found), found)

    def test_import_prohibido_entre_capas_nombra_fichero_linea_y_regla(self):
        found = self._findings(domain_source="\nfrom app.routes import x\n")
        self.assertEqual(found, ["HR-68: app/domain/x.py:2: import prohibido por la regla "
                                 "«core must not import web»"])

    def test_configuracion_del_linter_desfasada_del_documento_falla(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d)
            (repo / ".importlinter").write_text("[importlinter]\nroot_packages =\n    app\n", encoding="utf-8")
            found = CLI._check_architecture(repo, {"architecture": DECL})
        self.assertTrue(any(".importlinter" in f and "no coincide" in f for f in found), found)

    def test_stack_typescript_impone_las_reglas_con_dependency_cruiser(self):
        block = {"layers": [{"name": "web", "modules": ["src/web"]}, {"name": "core", "modules": ["src/core"]}],
                 "allowed": {"web": ["core"], "core": []}}
        doc = (DOC.replace("app/routes", "src/web").replace("app/domain", "src/core")
               .replace(json.dumps(BLOCK), json.dumps(block)))
        decl = {**DECL, "code_root": "src", "code_glob": "*.ts", "stack": "typescript",
                "linter_config": ".dependency-cruiser.cjs"}
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, doc=doc, modules=())
            _write(repo / ".dependency-cruiser.cjs", TS.render_depcruise(block))
            _write(repo / "src" / "web" / "view.ts", 'import { m } from "../core/model";\nexport const v = m;\n')
            _write(repo / "src" / "core" / "model.ts", "export const m = 1;\n")
            self.assertEqual(CLI._check_architecture(repo, {"architecture": decl}), [])
            _write(repo / "src" / "core" / "model.ts", 'import { v } from "../web/view";\nexport const m = v;\n')
            found = CLI._check_architecture(repo, {"architecture": decl})
        self.assertEqual(found, ["HR-68: src/core/model.ts:1: import prohibido por la regla "
                                 "«core-must-not-import-web»"])

    def test_job_ausente_del_workflow_falla(self):
        found = self._findings(workflow="jobs:\n  lint:\n    runs-on: x\n")
        self.assertTrue(any("HR-68" in f and "docs-architecture" in f and "job" in f for f in found), found)

    def test_job_fuera_de_required_jobs_falla(self):
        found = self._findings(policy={"required_jobs": ["unit"]})
        self.assertTrue(any("required_jobs" in f and "docs-architecture" in f for f in found), found)

    def test_sin_politica_de_required_jobs_falla_cerrado(self):
        found = self._findings(policy=None)
        self.assertTrue(any("required-jobs-policy.json" in f for f in found), found)


if __name__ == "__main__":
    unittest.main()
