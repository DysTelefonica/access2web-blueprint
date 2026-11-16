#!/usr/bin/env python3
# ci-pattern asset tests -- plantilla del workflow nocturno pre-release (DysTelefonica/team-skills#1627, slice 3)
"""``nightly-pipeline.template.yml``: contrato del workflow, leido solo con la stdlib.

No usa PyYAML (el catalogo no añade dependencias): separa los bloques de ``jobs:`` por
indentacion y compara las lineas exactas que importan. Comprueba que el gate necesita
exactamente los seis jobs requeridos y ejecuta el ``plan``, que el deploy solo corre con la
expresion exacta, y que el paso ``mig`` no escribe ningun valor (asi el gate retiene).
"""
import unittest
from pathlib import Path

ASSETS = Path(__file__).resolve().parents[1]
TEMPLATE = ASSETS / "nightly-pipeline.template.yml"
REQUIRED_JOBS = ["build", "e2e-full", "integration", "upgrade-path", "trivy", "security-deep"]
NEEDS_LINE = "needs: [" + ", ".join(REQUIRED_JOBS) + "]"
DEPLOY_IF = "if: ${{ inputs.dry_run != true && needs.gate.result == 'success' }}"


def job_blocks(text: str) -> dict:
    """Bloques ``  <job>:`` bajo ``jobs:``, cada uno como lista de lineas (sin indentar el bloque)."""
    lines = text.splitlines()
    start = lines.index("jobs:")
    blocks, current = {}, None
    for line in lines[start + 1:]:
        if line.startswith("  ") and not line.startswith("   ") and line.strip().endswith(":"):
            current = line.strip()[:-1]
            blocks[current] = []
        elif current is not None and line.strip():
            blocks[current].append(line)
    return blocks


def stripped(block: list) -> list:
    return [line.strip() for line in block]


class TemplateTest(unittest.TestCase):
    def setUp(self) -> None:
        self.text = TEMPLATE.read_text(encoding="utf-8")
        self.jobs = job_blocks(self.text)

    def test_reusable_and_dispatchable_with_a_dry_run_input(self) -> None:
        for trigger in ("schedule:", "workflow_dispatch:", "workflow_call:"):
            self.assertIn(trigger, self.text)
        self.assertEqual(self.text.count("dry_run:"), 2)
        self.assertEqual(self.text.count("type: boolean"), 2)

    def test_gate_needs_exactly_the_required_jobs_and_runs_the_plan(self) -> None:
        self.assertIn(NEEDS_LINE, stripped(self.jobs["gate"]))
        self.assertTrue(any("nightly_pipeline.py plan" in line for line in self.jobs["gate"]))

    def test_deploy_needs_the_gate_and_has_the_exact_condition(self) -> None:
        deploy = stripped(self.jobs["deploy"])
        self.assertIn("needs: gate", deploy)
        self.assertIn(DEPLOY_IF, deploy)

    def test_build_mig_step_writes_no_value(self) -> None:
        build = self.jobs["build"]
        starts = [i for i, line in enumerate(build) if line.strip() == "- id: mig"]
        self.assertEqual(len(starts), 1)
        step = []
        for line in build[starts[0] + 1:]:
            if line.strip().startswith("- "):
                break
            step.append(line)
        self.assertFalse(any("pending" in line for line in step), step)

    def test_every_placeholder_job_fails_closed(self) -> None:
        for job in REQUIRED_JOBS + ["deploy"]:
            with self.subTest(job=job):
                steps = [line.strip().removeprefix("- ") for line in self.jobs[job]]
                runs = [line for line in steps if line.startswith("run:")]
                self.assertTrue(any("PLACEHOLDER" in run and "exit 1" in run for run in runs), runs)


if __name__ == "__main__":
    unittest.main()
