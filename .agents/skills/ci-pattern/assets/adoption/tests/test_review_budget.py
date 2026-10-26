"""A1 y el mapa de orígenes: la fase 4 detecta al consumer que no lo cableó (#1159).

El presupuesto de revisión (A1) solo excluye el árbol vendorizado
(`.agents/skills/`) si el consumer pasa el mapa de orígenes al gate de contrato
de PR con `--origins-file`. Sin el mapa el gate cuenta todo y una PR de
propagación del catálogo falla con un A1 que no es trabajo de quien abre la PR
(«527 > 400» sobre un diff que solo trae el catálogo; #1133, cadete#643). La
fase 4 (G4.2) lo exige: con `budget` activo en la política del gate y el mapa
sin cablear, el hallazgo dice qué cablear. `pr-size` mide A1 por ruta y NO
necesita mapa: quien solo lo ejecuta no recibe el hallazgo.
"""

import importlib.machinery
import importlib.util
import json
import tempfile
import unittest
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[3]
CLI_PATH = SKILL_ROOT / "assets" / "bin" / "ci-pattern"


def _load_cli():
    spec = importlib.util.spec_from_loader(
        "ci_pattern_cli_review_budget",
        importlib.machinery.SourceFileLoader("ci_pattern_cli_review_budget", str(CLI_PATH)),
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


CLI = _load_cli()

# El cableado canonico: un paso que genera el mapa y el gate que lo recibe.
GENERADOR = """      - name: Compute the origins file
        run: |
          python3 "$PATTERN_DIR/assets/bin/origins_file.py" \\
            --base "${BASE_SHA}" \\
            --head "${HEAD_SHA}" \\
            --output .ci-pr-origins
"""
GATE = """      - name: PR contract gate
        run: |
          # el gate assets/pr-contract/check_pr_contract.py no ejecuta git: el
          # workflow le pasa los datos del rango (este comentario no es una invocacion)
          python3 -I "$PATTERN_DIR/assets/pr-contract/check_pr_contract.py" \\
            --event-file     "$GITHUB_EVENT_PATH" \\
            --policy-file    .github/pr-contract.policy.json \\
            --numstat-file   .ci-pr-numstat \\
            --origins-file   .ci-pr-origins
"""
WORKFLOW_CABLEADO = "jobs:\n  pr-contract:\n    steps:\n" + GENERADOR + GATE
# El mismo gate sin el mapa: no recibe `--origins-file`.
WORKFLOW_SIN_MAPA = WORKFLOW_CABLEADO.replace(
    "            --origins-file   .ci-pr-origins\n", "")
# El gate recibe el mapa pero ningun paso lo genera: el gate lo leera vacio.
WORKFLOW_MAPA_SIN_GENERAR = WORKFLOW_CABLEADO.replace(GENERADOR, "")
# Un consumer que mide A1 con el gate `pr-size`: no necesita mapa.
WORKFLOW_SOLO_PR_SIZE = """jobs:
  pr-gate:
    steps:
      - run: python .agents/skills/ci-pattern/assets/pr-size/check_pr_size.py --event "$E"
"""

# La plantilla canonica lee la politica de la BASE (`base/`, un checkout en CI). Fuera
# de CI no existe `base/`: la copia versionada del repo es la misma politica (#1313).
WORKFLOW_BASE = WORKFLOW_CABLEADO.replace("--policy-file    .github/pr-contract.policy.json",
                                          "--policy-file    base/.github/pr-contract.policy.json")
WORKFLOW_BASE_SIN_MAPA = WORKFLOW_SIN_MAPA.replace(
    "--policy-file    .github/pr-contract.policy.json",
    "--policy-file    base/.github/pr-contract.policy.json")

POLICY_BUDGET_ACTIVO = {"checks": {"budget": True, "commits": True}}
POLICY_BUDGET_APAGADO = {"checks": {"budget": False, "commits": True}}
POLICY_SIN_CHECKS = {"review_budget_lines": 400}


def _write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _repo(root, workflow=WORKFLOW_CABLEADO, policy=POLICY_BUDGET_ACTIVO, policy_name=None):
    repo = Path(root)
    _write(repo / "app" / "users.py", "x = 1\n")
    if workflow is not None:
        _write(repo / ".github" / "workflows" / "pr-contract.yml", workflow)
    if policy is not None:
        name = policy_name or ".github/pr-contract.policy.json"
        _write(repo / name, policy if isinstance(policy, str) else json.dumps(policy))
    return repo


class ReviewBudgetTests(unittest.TestCase):
    """El detector de la fase 4 sobre el arbol del consumer."""

    def _findings(self, **kw):
        with tempfile.TemporaryDirectory() as d:
            return CLI._check_review_budget(_repo(d, **kw))

    def test_mapa_cableado_da_verde(self):
        # El canon: el paso genera el mapa y el gate lo recibe.
        self.assertEqual(self._findings(), [])

    def test_a1_activo_sin_mapa_da_hallazgo_que_dice_que_cablear(self):
        # Un hallazgo, y solo uno: el comentario que nombra el gate no cuenta
        # como invocacion.
        found = self._findings(workflow=WORKFLOW_SIN_MAPA)
        self.assertEqual(len(found), 1, found)
        message = " ".join(found)
        for expected in ("--origins-file", "origins_file.py", "pr-contract.yml.tmpl",
                         "check_pr_contract.py"):
            self.assertIn(expected, message)

    def test_a1_activo_con_mapa_sin_generar_da_hallazgo(self):
        # El gate recibe un fichero que nadie escribe: lo leera vacio y
        # fallara cerrado al medir el presupuesto.
        found = self._findings(workflow=WORKFLOW_MAPA_SIN_GENERAR)
        self.assertTrue(found, "el mapa que nadie genera debe ser hallazgo")
        self.assertIn(".ci-pr-origins", " ".join(found))

    def test_budget_apagado_no_exige_mapa(self):
        # `pr-size` ya mide el presupuesto: el mapa no hace falta.
        self.assertEqual(self._findings(workflow=WORKFLOW_SIN_MAPA,
                                        policy=POLICY_BUDGET_APAGADO), [])

    def test_sin_checks_todo_activo_exige_mapa(self):
        # Sin `checks` estan todas las comprobaciones activas, `budget` entre
        # ellas: el mapa es obligatorio.
        found = self._findings(workflow=WORKFLOW_SIN_MAPA, policy=POLICY_SIN_CHECKS)
        self.assertTrue(found, "sin `checks` la comprobacion budget esta activa")

    def test_consumer_solo_con_pr_size_no_exige_mapa(self):
        # BP y APAP_WEB miden A1 con `pr-size`: no hay gate de contrato de PR
        # que cablear.
        self.assertEqual(self._findings(workflow=WORKFLOW_SOLO_PR_SIZE), [])

    def test_politica_ilegible_no_es_verde(self):
        # Sin politica legible no hay medicion: un gate que no puede medir no
        # sale en verde (HR-3).
        found = self._findings(policy="{ esto no es json")
        self.assertTrue(found, "una politica ilegible no puede dar verde")
        self.assertIn("ilegible", " ".join(found))

    def test_el_consumer_declara_la_politica_que_usa_el_gate(self):
        # La politica que cuenta es la que el gate recibe con
        # `--policy-file`, no la que el consumer tenga por costumbre.
        found = self._findings(workflow=WORKFLOW_SIN_MAPA,
                               policy=POLICY_BUDGET_ACTIVO,
                               policy_name=".github/otra.policy.json")
        self.assertTrue(found, "el gate usa .github/pr-contract.policy.json: esa "
                               "politica no existe y no hay medicion")


class ReviewBudgetBasePolicyTests(unittest.TestCase):
    """La politica que la plantilla lee de `base/` (DysTelefonica/team-skills#1313, G4)."""

    def _findings(self, workflow, repo_policy=POLICY_BUDGET_ACTIVO, base_policy=None):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, workflow=workflow, policy=repo_policy)
            if base_policy is not None:
                _write(repo / "base" / ".github" / "pr-contract.policy.json",
                       json.dumps(base_policy))
            return CLI._check_review_budget(repo)

    def test_sin_base_checkout_se_lee_la_copia_del_repo(self):
        # Fuera de CI no hay `base/`: la politica versionada en el repo es la que cuenta.
        self.assertEqual(self._findings(WORKFLOW_BASE), [])

    def test_sin_base_ni_copia_sigue_siendo_ilegible(self):
        # Fail closed: ninguna de las dos copias existe, no hay medicion.
        found = self._findings(WORKFLOW_BASE, repo_policy=None)
        self.assertTrue(found, "sin ninguna politica no puede dar verde")
        self.assertIn("ilegible", " ".join(found))

    def test_el_checkout_de_base_gana_a_la_copia_del_repo(self):
        # La copia de la base es la de confianza: con `budget` apagado en la base,
        # la copia del repo (encendida) no exige el mapa.
        self.assertEqual(self._findings(WORKFLOW_BASE_SIN_MAPA,
                                        repo_policy=POLICY_BUDGET_ACTIVO,
                                        base_policy=POLICY_BUDGET_APAGADO), [])


class ReviewBudgetPhaseTests(unittest.TestCase):
    """La comprobacion vive en la fase 4 del contrato de adopcion."""

    def test_la_fase_4_llama_al_comprobador(self):
        with tempfile.TemporaryDirectory() as d:
            repo = _repo(d, workflow=WORKFLOW_SIN_MAPA)
            contract = {"schema_version": 1, "phases": {"4": {"gates": []}}}
            findings = CLI._phase_findings(repo, 4, {"gates": []}, contract)
        self.assertTrue(any("--origins-file" in f for f in findings), findings)


if __name__ == "__main__":
    unittest.main()
