#!/usr/bin/env python3
# ci-pattern asset tests — deploy-evidence gate (HR-10; team-skills#253)
"""Suite del gate de evidencia de deploy (HR-10): los workflows de deploy o
batería declarados en la política publican su evidencia como estado/check de
commit sobre el SHA desplegado — nunca sobre la rama por defecto ni colgada
de una variable global. Ejecución: `python3 test_deploy_evidence.py`.
Solo stdlib, sin red."""

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_deploy_evidence.py"
SKILL_ROOT = Path(__file__).resolve().parents[3]
REPO_ROOT = SKILL_ROOT.parents[2]

CLEAN_DEPLOY = """jobs:
  deploy:
    runs-on: ubuntu-24.04
    steps:
      - name: Publica evidencia sobre el SHA desplegado
        run: |
          gh api repos/$GITHUB_REPOSITORY/statuses/${{ github.sha }} \\
            --field state=success
"""

BRANCH_EVIDENCE = """jobs:
  deploy:
    runs-on: ubuntu-24.04
    steps:
      - name: Evidencia sobre la rama por defecto
        run: |
          gh api repos/$GITHUB_REPOSITORY/statuses/${{ github.ref }} \\
            --field state=success
"""

GLOBAL_ENV_EVIDENCE = """env:
  DEPLOYED_SHA: pending
jobs:
  deploy:
    runs-on: ubuntu-24.04
    steps:
      - name: Evidencia sobre variable global
        run: |
          gh api repos/$GITHUB_REPOSITORY/statuses/$DEPLOYED_SHA \\
            --field state=success
"""

NO_EVIDENCE = """jobs:
  deploy:
    runs-on: ubuntu-24.04
    steps:
      - name: Deploy sin evidencia
        run: echo "desplegado"
"""


def build_repo(tmp, workflow, policy=None):
    repo = Path(tmp) / "repo"
    wf = repo / ".github" / "workflows"
    wf.mkdir(parents=True)
    (wf / "deploy.yml").write_text(workflow, encoding="utf-8")
    pol = repo / ".github" / "deploy-evidence-policy.json"
    pol.write_text(json.dumps(policy if policy is not None else {
        "evidence_jobs": [{"workflow": ".github/workflows/deploy.yml", "job": "deploy"}]
    }), encoding="utf-8")
    return repo, pol


def run_gate(repo, policy_path):
    return subprocess.run([sys.executable, str(GATE), "--policy", str(policy_path),
                           "--repo", str(repo)],
                          capture_output=True, text=True, check=False)


class DeployEvidenceGateTests(unittest.TestCase):

    def test_evidence_on_deployed_sha_exits_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, pol = build_repo(tmp, CLEAN_DEPLOY)
            proc = run_gate(repo, pol)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("DEPLOY EVIDENCE OK", proc.stdout)

    def test_evidence_on_default_branch_is_a_finding(self):
        """Fixture de violación (HR-10): evidencia publicada sobre la rama
        por defecto (github.ref) en vez del SHA desplegado."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, pol = build_repo(tmp, BRANCH_EVIDENCE)
            proc = run_gate(repo, pol)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("rama por defecto", proc.stdout)

    def test_global_env_evidence_is_a_finding(self):
        """Fixture de violación (HR-10): el SHA cuelga de una variable
        global declarada a nivel de workflow."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, pol = build_repo(tmp, GLOBAL_ENV_EVIDENCE)
            proc = run_gate(repo, pol)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("variable global", proc.stdout)

    def test_missing_evidence_publication_is_a_finding(self):
        """Un job declarado como publicador de evidencia que no publica
        nada sobre commits es hallazgo."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, pol = build_repo(tmp, NO_EVIDENCE)
            proc = run_gate(repo, pol)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("no publica", proc.stdout)

    def test_declared_job_absent_fails_closed(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, pol = build_repo(tmp, CLEAN_DEPLOY)
            pol.write_text(json.dumps({
                "evidence_jobs": [{"workflow": ".github/workflows/deploy.yml",
                                   "job": "no-existe"}]}), encoding="utf-8")
            proc = run_gate(repo, pol)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("no existe", proc.stdout)

    def test_malformed_policy_fails_closed(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, pol = build_repo(tmp, CLEAN_DEPLOY)
            pol.write_text("{not json", encoding="utf-8")
            proc = run_gate(repo, pol)
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("fail-closed", proc.stderr)

    def test_empty_policy_is_a_declared_state_not_vacuous(self):
        """Un repo sin deploys declara la lista vacía: exit 0 con la línea
        explícita (estado declarado, no una medición vacía silenciosa)."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, pol = build_repo(tmp, CLEAN_DEPLOY, policy={"evidence_jobs": []})
            proc = run_gate(repo, pol)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("sin jobs de evidencia declarados", proc.stdout)


    def test_job_block_stops_at_next_sibling_job(self):
        """Hallazgo 1: el bloque del job declarado no se come los jobs
        siguientes; un deploy sin evidencia no hereda la del vecino."""
        wf = ("jobs:\n"
              "  deploy:\n"
              "    runs-on: ubuntu-24.04\n"
              "    steps:\n"
              "      - run: echo desplegado\n"
              "  other:\n"
              "    runs-on: ubuntu-24.04\n"
              "    steps:\n"
              "      - run: gh api repos/$GITHUB_REPOSITORY/statuses/${{ github.sha }}\n")
        with tempfile.TemporaryDirectory() as tmp:
            repo, pol = build_repo(tmp, wf)
            proc = run_gate(repo, pol)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("no publica evidencia sobre commits", proc.stdout)
            self.assertNotIn("DEPLOY EVIDENCE OK", proc.stdout)

    def test_needs_output_without_sha_is_not_sha_evidence(self):
        """Hallazgo 2: needs.build.outputs.branch NO es un SHA; el nombre de
        la salida debe contener sha (como dice el docstring del gate)."""
        wf = ("jobs:\n"
              "  deploy:\n"
              "    steps:\n"
              "      - run: gh api repos/$GITHUB_REPOSITORY/statuses/${{ needs.build.outputs.branch }}\n")
        with tempfile.TemporaryDirectory() as tmp:
            repo, pol = build_repo(tmp, wf)
            proc = run_gate(repo, pol)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("sin un SHA explícito", proc.stdout)

    def test_global_env_used_for_repo_not_sha_is_clean(self):
        """Hallazgo 3: la variable global en la posición del REPO no es
        hallazgo; solo cuenta si ocupa la posición del SHA tras statuses/
        o en head_sha=."""
        wf = ("env:\n"
              "  GH_REPO: ardelperal/APAP_WEB\n"
              "jobs:\n"
              "  deploy:\n"
              "    steps:\n"
              "      - run: gh api repos/$GH_REPO/statuses/${{ github.sha }} --field state=success\n")
        with tempfile.TemporaryDirectory() as tmp:
            repo, pol = build_repo(tmp, wf)
            proc = run_gate(repo, pol)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("DEPLOY EVIDENCE OK", proc.stdout)


class DeployEvidenceDogfoodingTests(unittest.TestCase):
    """Dogfooding sobre el árbol real de team-skills (no saltable)."""

    def test_team_skills_declares_no_evidence_jobs_and_gate_reports_it(self):
        policy = (REPO_ROOT / ".github" / "deploy-evidence-policy.json")
        self.assertTrue(policy.is_file(),
                        "team-skills debe declarar su política (aunque sea vacía)")
        proc = run([sys.executable, str(GATE), "--policy", str(policy),
                    "--repo", str(REPO_ROOT)])
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("sin jobs de evidencia declarados", proc.stdout)


def run(*argv):
    if len(argv) == 1 and isinstance(argv[0], (list, tuple)):
        argv = tuple(argv[0])
    return subprocess.run([str(a) for a in argv], capture_output=True,
                          text=True, check=False)


if __name__ == "__main__":
    unittest.main(verbosity=2)
