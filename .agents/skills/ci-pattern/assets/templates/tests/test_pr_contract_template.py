#!/usr/bin/env python3
# ci-pattern asset tests — pr-contract template, model A: canon + gate (#1261)
"""The pr-contract template runs without a secret in the gate job.

The catalog token is resolved by the `canon` job, which loads the action from
the base and publishes the pattern as an artifact. The gate job consumes that
artifact and holds no secret, so it can run the code of the PR.

Executes offline; no consumer and no network.
"""
import re
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent.parent
TEMPLATE = ASSET / "templates" / ".github" / "workflows" / "pr-contract.yml.tmpl"


class PrContractTemplateTests(unittest.TestCase):
    def setUp(self):
        self.text = TEMPLATE.read_text(encoding="utf-8")

    def test_the_template_declares_no_secret_input(self):
        """El `workflow_call` no exige ningún secreto a quien lo llama."""
        self.assertNotRegex(self.text, r"(?m)^\s+secrets:\s*$")

    def test_the_gate_job_holds_no_secret(self):
        """El job que mide el PR no puede tener el token del catálogo."""
        gate = self.text[self.text.index("  pr-contract:"):]
        self.assertNotIn("secrets.", gate)

    def test_the_gate_job_needs_the_canon_job(self):
        self.assertIn("needs: canon", self.text)

    def test_the_gate_runs_from_the_published_canon(self):
        self.assertRegex(self.text, r"(?m)^\s+PATTERN_DIR: \.canon$")
        self.assertIn('"$PATTERN_DIR/assets/pr-contract/check_pr_contract.py"', self.text)
        self.assertIn('"$PATTERN_DIR/assets/bin/origins_file.py"', self.text)

    def test_checkout_fetches_the_base_history_for_the_range(self):
        """`git diff BASE...HEAD` necesita el historial de la base."""
        self.assertRegex(self.text, r"fetch-depth:\s*0")

    def test_every_checkout_pins_the_same_sha(self):
        """Todos los checkouts fijan el MISMO SHA.

        Un SHA distinto no existe en el action: el checkout falla con HTTP 422
        y el job del gate no llega a correr. El detector de simbolos muertos no
        lo ve (es una cadena, no un simbolo), asi que este test es el que lo
        atrapa.
        """
        pins = re.findall(r"actions/checkout@([0-9a-zA-Z]+)", self.text)
        self.assertGreaterEqual(len(pins), 2, "el template debe tener al menos dos checkouts")
        distintos = sorted(set(pins))
        self.assertEqual(len(distintos), 1,
                         f"los checkouts no fijan el mismo SHA: {distintos}")
        for pin in pins:
            self.assertRegex(pin, r"^[0-9a-f]{40}$", f"SHA de checkout invalido: {pin}")

    def test_the_action_comes_from_the_base_checkout(self):
        """El checkout del job `canon` es de la base: su acción es la de la base."""
        canon = self.text[self.text.index("  canon:"):self.text.index("  pr-contract:")]
        self.assertIn("uses: ./.github/actions/pattern-at-pin", canon)
        self.assertRegex(canon, r"ref: \$\{\{ github\.event\.pull_request\.base\.sha \|\| github\.sha \}\}")

    def test_the_workflow_runs_on_pull_request_target(self):
        """En `pull_request` el workflow es el del PR: el secret no puede estar ahí."""
        self.assertRegex(self.text, r"(?m)^  pull_request_target:")
        self.assertIsNone(re.search(r"(?m)^  pull_request:", self.text))

    def test_the_pr_head_enters_as_data(self):
        """La cabeza del PR se checkoutea aparte y el gate la lee, no la ejecuta."""
        gate = self.text[self.text.index("  pr-contract:"):]
        self.assertIn("path: pr-head", gate)
        self.assertIn("git -C pr-head diff", gate)

    def test_the_verdict_is_published_on_the_pr_head(self):
        """Un run de `pull_request_target` se atribuye a la base: el veredicto va como estado."""
        self.assertIn("check-runs", self.text)
        self.assertIn("checks: write", self.text)

    def test_the_published_context_is_the_name_the_host_contract_requires(self):
        """El veredicto se publica con el nombre del check run.

        Los contratos de la flota exigen el contexto `pr-contract / pr-contract`
        (cadete `.github/host-contract.json`, P28). Con otro nombre, el check
        requerido no aparece en la cabeza del PR y el gate no bloquea el merge.
        """
        self.assertRegex(self.text, r'(?m)^\s+CONTEXT: "pr-contract / pr-contract"$')

    def test_the_published_context_matches_the_workflow_and_job_names(self):
        """GitHub nombra el job `<workflow> / <job>`: el contexto debe coincidir."""
        workflow = re.search(r"(?m)^name: (\S+)$", self.text)
        job = re.search(r"(?m)^  pr-contract:\n    name: (\S+)$", self.text)
        context = re.search(r'(?m)^\s+CONTEXT:\s*"?([^"\n]+?)"?\s*$', self.text)
        self.assertIsNotNone(workflow, "el workflow debe declarar su nombre")
        self.assertIsNotNone(job, "el job del gate debe declarar su nombre")
        self.assertIsNotNone(context, "el workflow debe publicar el veredicto con un CONTEXT")
        self.assertEqual(context.group(1).strip(), f"{workflow.group(1)} / {job.group(1)}")


    def test_no_configuration_is_read_from_the_pr_head(self):
        """Un PR no puede reescribir la política que lo mide."""
        for cfg in (".github/ci-pattern-adoption.json", ".github/pr-contract.policy.json",
                    ".github/host-contract.json", "ci-pattern.yaml"):
            with self.subTest(configuracion=cfg):
                self.assertNotIn(f"pr-head/{cfg}", self.text)
                self.assertNotRegex(self.text, rf"pr-head/[^\s\"']*{re.escape(cfg.split('/')[-1])}")

    def test_the_policy_comes_from_the_base(self):
        self.assertIn("--policy-file    base/.github/pr-contract.policy.json", self.text)
        self.assertIn("path: base", self.text)
        self.assertRegex(self.text, r"ref: \$\{\{ github\.event\.pull_request\.base\.sha \}\}")

    def test_the_base_checkout_leaves_no_credentials(self):
        base = self.text[self.text.index("path: base"):]
        self.assertIn("persist-credentials: false", base[:400])



if __name__ == "__main__":
    unittest.main()
