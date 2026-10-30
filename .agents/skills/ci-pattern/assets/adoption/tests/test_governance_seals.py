"""HR-72 (#1027): la fase 3 —la identidad del gobierno: los sellos de
`.governance-manifest.json` y el cableado de `P23_required_jobs` (HR-46)— se
mide EN UN PULL REQUEST, no solo fuera de el.

Con la declaracion `governance_seals` en el contrato de adopcion, el canon exige
un job de PR que la ejecute, y la fase 3 comprueba que ese job existe en un
workflow y figura en `required_jobs` de la politica del consumer: la unica via de
enforcement que funciona en repos privados con plan Free, donde el host no ofrece
ni proteccion de rama ni required checks (host_capabilities en unavailable y
`required_checks: runner-enforced`). Sin declaracion avisa a stderr y sale 0, como
en HR-55 y HR-63.
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


def _load_cli():
    spec = importlib.util.spec_from_loader(
        "ci_pattern_cli_governance_seals",
        importlib.machinery.SourceFileLoader("ci_pattern_cli_governance_seals", str(CLI_PATH)),
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


CLI = _load_cli()
DECL = {"job": "governance-seals"}
# El job vive en `ci.yml` junto al resto de gates y el agregador `required` lo
# exige por `needs:`: mismo camino que cualquier otro gate del patron.
WORKFLOW = "jobs:\n  governance-seals:\n    runs-on: ubuntu-latest\n"
POLICY = {"required_jobs": ["unit", "governance-seals"]}


def _write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _repo(root, decl_overrides=None, workflow=WORKFLOW, policy=POLICY):
    repo = Path(root)
    _write(repo / "app" / "users.py", "x = 1\n")
    if workflow is not None:
        _write(repo / ".github" / "workflows" / "ci.yml", workflow)
    if policy is not None:
        _write(repo / ".github" / "required-jobs-policy.json", policy
               if isinstance(policy, str) else json.dumps(policy))
    decl = {**DECL, **(decl_overrides or {})} if decl_overrides is not False else None
    return repo, ({"governance_seals": decl} if decl is not None else {})


class GovernanceSealsTests(unittest.TestCase):
    def _findings(self, **kw):
        with tempfile.TemporaryDirectory() as d:
            repo, contract = _repo(d, **kw)
            return CLI._check_governance_seals(repo, contract)

    def test_conforme_da_verde(self):
        self.assertEqual(self._findings(), [])

    def test_sin_declaracion_avisa_por_stderr_y_sale_verde(self):
        # Fail-open: un consumer que no declara el job conserva su adopcion
        # previa. El que lo declare queda obligado (HR-3 al reves: aqui la
        # ausencia es la no-exigencia, explicita).
        with tempfile.TemporaryDirectory() as d:
            buf = io.StringIO()
            with contextlib.redirect_stderr(buf):
                self.assertEqual(CLI._check_governance_seals(Path(d), {}), [])
            self.assertIn("HR-72: governance_seals no declarado", buf.getvalue())

    def test_declaracion_incompleta_falla_nombrando_la_clave(self):
        # Sin `job` no hay nada que medir: rojo nombrando la clave, nunca
        # un verde con la declaracion a medias (HR-3).
        found = self._findings(decl_overrides={"job": "  "})
        self.assertTrue(any("job" in f for f in found), found)

    def test_declaracion_desconocida_falla_cerrado(self):
        # Clave que el canon no conoce: hallazgo, no silencio (esquema cerrado).
        found = self._findings(decl_overrides={"jobs": "gobierno"})
        self.assertTrue(any("jobs" in f for f in found), found)

    def test_job_que_ningun_workflow_define_falla(self):
        found = self._findings(workflow="jobs:\n  unit:\n    runs-on: ubuntu-latest\n")
        self.assertTrue(any("governance-seals" in f and "ningun workflow" in f
                            for f in found), found)

    def test_job_que_no_figura_en_required_jobs_falla(self):
        # Sin figurar en `required_jobs` el job corre pero NADIE lo exige: en un
        # repo privado con plan Free es la unica via de enforcement que queda,
        # asi que la ausencia es el hueco que esta HR cierra.
        found = self._findings(policy={"required_jobs": ["unit"]})
        self.assertTrue(any("required_jobs" in f for f in found), found)

    def test_politica_ilegible_falla_cerrado(self):
        # Sin politica legible no hay medicion; un gate que no puede medir
        # nunca sale en verde (HR-3).
        found = self._findings(policy="{ esto no es json")
        self.assertTrue(any("ilegible" in f for f in found), found)


if __name__ == "__main__":
    unittest.main()
