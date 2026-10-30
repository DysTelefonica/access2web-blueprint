#!/usr/bin/env python3
# ci-pattern asset tests — pr-contract template: one job, pattern from the base (#1396)
"""The pr-contract template is ONE job that holds no secret and uploads nothing.

The pattern comes from the vendored copy in the checked-out BASE
(`base/.agents/skills/ci-pattern`), as the consumers' `ci.yml` jobs already do:
no catalog token, no `canon` job, no artifact that a billing quota can block
(DysTelefonica/cadete#752). The verdict is mirrored as a check run on the PR head,
because the native check run of a `pull_request_target` run lands on the base
commit and the governed merge reads the head; the reason is documented (#1396).

Executes offline; no consumer and no network.
"""
import json
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent.parent
TEMPLATE = ASSET / "templates" / ".github" / "workflows" / "pr-contract.yml.tmpl"
POLICY_GATE = ASSET / "workflow-policy" / "check_workflow_policy.py"


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

    def test_the_template_has_a_single_job(self):
        """Sin job `canon` ni `needs:`: el gate no espera a nadie (#1396)."""
        self.assertNotIn("  canon:", self.text)
        self.assertNotRegex(self.text, r"(?m)^\s+needs:")
        self.assertEqual(re.findall(r"(?m)^  ([a-z0-9-]+):$", self.text[self.text.index("jobs:"):]),
                         ["pr-contract"])

    def test_no_artifact_is_uploaded_or_downloaded(self):
        """Una cuota de facturación no puede saltar el check requerido (cadete#752)."""
        self.assertNotIn("upload-artifact", self.text)
        self.assertNotIn("download-artifact", self.text)

    def test_the_template_reads_no_catalog_token(self):
        self.assertNotIn("TEAM_SKILLS", self.text)
        self.assertNotIn("secrets.", self.text)
        self.assertNotIn("pattern-at-pin", self.text)

    def test_the_gate_runs_from_the_pattern_vendored_in_the_base(self):
        """El patrón es la copia vendorizada de la BASE: el PR no elige el código que lo mide."""
        self.assertRegex(self.text, r"(?m)^\s+PATTERN_DIR: base/\.agents/skills/ci-pattern$")
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

    def test_the_workflow_runs_on_pull_request_target(self):
        """En `pull_request` el workflow es el del PR: el secret no puede estar ahí."""
        self.assertRegex(self.text, r"(?m)^  pull_request_target:")
        self.assertIsNone(re.search(r"(?m)^  pull_request:", self.text))

    def test_the_pr_head_enters_as_data(self):
        """La cabeza del PR se checkoutea aparte y el gate la lee, no la ejecuta."""
        gate = self.text[self.text.index("  pr-contract:"):]
        self.assertIn("path: pr-head", gate)
        self.assertIn("git -C pr-head diff", gate)

    def test_the_verdict_is_mirrored_on_the_pr_head(self):
        """El check run nativo de `pull_request_target` queda en la base y el merge
        gobernado lee la cabeza: el espejo se queda (#1396, S9)."""
        self.assertIn("check-runs", self.text)
        self.assertIn("checks: write", self.text)
        self.assertIn("HEAD_SHA: ${{ github.event.pull_request.head.sha }}", self.text)

    def test_the_reason_for_the_duplicate_check_is_documented(self):
        """S9: cada gate sale una vez, o el motivo del duplicado está escrito."""
        self.assertIn("Por que el gate sale dos veces en los checks del PR (#1396, S9)", self.text)
        self.assertIn("repos/{repo}/commits/{head_sha}/check-runs", self.text)

    def test_the_published_context_matches_the_workflow_and_job_names(self):
        """GitHub nombra el job `<workflow> / <job>`: el contexto debe coincidir (P28)."""
        workflow = re.search(r"(?m)^name: (\S+)$", self.text)
        job = re.search(r"(?m)^  pr-contract:\n    name: (\S+)$", self.text)
        context = re.search(r'(?m)^\s+CONTEXT:\s*"?([^"\n]+?)"?\s*$', self.text)
        self.assertIsNotNone(workflow, "el workflow debe declarar su nombre")
        self.assertIsNotNone(job, "el job del gate debe declarar su nombre")
        self.assertIsNotNone(context, "el workflow debe publicar el veredicto con un CONTEXT")
        self.assertEqual(context.group(1).strip(), f"{workflow.group(1)} / {job.group(1)}")

    def test_the_policy_gate_accepts_the_required_context_of_the_contract(self):
        """El contrato exige `pr-contract / pr-contract`: lo publica el espejo (#1313)."""
        code, out = self._policy_gate("pr-contract / pr-contract")
        self.assertEqual(code, 0, out)

    def _policy_gate(self, required_check):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            workflow, policy, contract = (root / "pr-contract.yml", root / "policy.json",
                                          root / "host-contract.json")
            workflow.write_text(self.text, encoding="utf-8")
            policy.write_text(json.dumps({
                "required_jobs": ["pr-contract"],
                "events": {"pull_request_target": {"accepted_skips": []}},
                "state_dependent_gates": []}), encoding="utf-8")
            contract.write_text(json.dumps({"contract_version": 1, "required_checks": [
                {"name": required_check, "class": "runner-enforced"}]}), encoding="utf-8")
            proc = subprocess.run([sys.executable, str(POLICY_GATE), "--workflow", str(workflow),
                                   "--required-jobs", str(policy), "--base-branch", "main",
                                   "--repo-root", str(root), "--host-contract", str(contract)],
                                  capture_output=True, text=True, check=False)
        return proc.returncode, proc.stdout + proc.stderr

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
