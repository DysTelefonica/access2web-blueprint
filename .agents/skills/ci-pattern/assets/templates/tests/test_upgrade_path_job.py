#!/usr/bin/env python3
# ci-pattern asset tests — forma canónica del job upgrade-path (HR-92, #1561)
"""El job `upgrade-path` que publica el porting-guide (G4.7) es el que los gates aceptan.

Como en `test_governance_seals_job.py`, la forma canónica vive en el documento y
el test la extrae y la mide: la prosa y el gate no pueden divergir.

1. Pegado en el `ci.yml` de un consumer, con su entrada en `required_jobs`, pasa
   `check_workflow_policy.py` del propio canon.
2. Prueba lo que dice probar: el esquema de partida sale de la última release
   (no de una base vacía), el arranque del cambio corre dos veces sobre la misma
   base y la imagen de Postgres va fijada por digest (HR-78).

Ejecuta sin red y sin consumer.
"""
import importlib
import json
import re
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SKILL_ROOT = ASSET.parents[1]
REPO = ASSET.parents[4]
PORTING_GUIDE = SKILL_ROOT / "references" / "porting-guide.md"
GATE = SKILL_ROOT / "assets" / "workflow-policy" / "check_workflow_policy.py"

if str(ASSET.parent / "workflow-policy") not in sys.path:
    sys.path.insert(0, str(ASSET.parent / "workflow-policy"))
policy_gate = importlib.import_module("check_workflow_policy")

JOB = "upgrade-path"
# El `ci.yml` mínimo de un consumer: el fragmento entra como un job más y el
# agregador `required` lo espera en su `needs`, que es como lo cablea quien adopta.
SKELETON = """name: ci
on:
  pull_request:
concurrency:
  group: ci-${{{{ github.ref }}}}
  cancel-in-progress: false
jobs:
{job}
  required:
    needs: [upgrade-path]
    runs-on: ubuntu-24.04
    timeout-minutes: 5
    steps:
      - run: test "${{{{ needs.upgrade-path.result }}}}" = success
"""


def canonical_job() -> str:
    """El job que publica el porting-guide, con la sangría de `jobs:`."""
    text = PORTING_GUIDE.read_text(encoding="utf-8")
    for block in re.findall(r"```yaml\n(.*?)```", text, re.S):
        if re.search(rf"^\s*{JOB}:\s*$", block, re.M):
            return textwrap.indent(textwrap.dedent(block).rstrip("\n"), "  ")
    raise AssertionError(f"el porting-guide ya no publica el job `{JOB}` (G4.7, HR-92)")


def consumer_workflow() -> str:
    return SKELETON.format(job=canonical_job())


def job_tree() -> dict:
    return policy_gate.parse_workflow(consumer_workflow(), "ci.yml")["jobs"][JOB]


def commands() -> str:
    return "\n".join(str(step.get("run") or "") for step in job_tree()["steps"])


class UpgradePathJobTests(unittest.TestCase):
    def test_fragment_passes_the_workflow_policy_gate_as_a_required_job(self):
        with tempfile.TemporaryDirectory() as tmp:
            rendered = Path(tmp) / "ci.yml"
            rendered.write_text(consumer_workflow(), encoding="utf-8")
            policy = Path(tmp) / "policy.json"
            policy.write_text(json.dumps({"required_jobs": [JOB, "required"]}), encoding="utf-8")
            proc = subprocess.run([sys.executable, str(GATE), "--workflow", str(rendered),
                                   "--required-jobs", str(policy), "--base-branch", "main"],
                                  capture_output=True, text=True, cwd=REPO)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("WORKFLOW-POLICY OK", proc.stdout)

    def test_the_starting_schema_is_the_last_release_not_an_empty_database(self):
        self.assertIn("git describe --tags --abbrev=0", commands())
        checkouts = [s for s in job_tree()["steps"] if str(s.get("uses", "")).startswith("actions/checkout@")]
        self.assertEqual(checkouts[0]["with"]["fetch-depth"], "0", "sin los tags no hay última release")
        self.assertEqual(checkouts[1]["with"]["ref"], "${{ steps.release.outputs.sha }}")

    def test_the_change_boots_twice_over_the_same_database(self):
        self.assertEqual(job_tree()["env"]["UPGRADE_BOOTS"], "2")
        self.assertIn("UPGRADE_BOOT_COMMAND", job_tree()["env"])
        self.assertIn("(cd release && $UPGRADE_BOOT_COMMAND)", commands())
        self.assertIn('for boot in $(seq "$UPGRADE_BOOTS")', commands())
        self.assertNotIn("dropdb", commands())

    def test_the_database_image_is_pinned_by_digest(self):
        image = job_tree()["services"]["postgres"]["image"]
        self.assertRegex(image, r"@sha256:")

    def test_the_job_does_not_swallow_the_verdict(self):
        self.assertNotIn("|| true", commands())
        self.assertNotIn("continue-on-error", str(job_tree()))


if __name__ == "__main__":
    unittest.main(verbosity=2)
