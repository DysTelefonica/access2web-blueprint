#!/usr/bin/env python3
"""Suite del liveness gate (F3, #396, HR-3/HR-21)."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_gate_liveness.py"
SUBP = subprocess.run


def _r(cmd, cwd, **kw):
    return SUBP(cmd, cwd=str(cwd), capture_output=True, text=True, **kw)


def _mk_repo(workflows: dict[str, str]) -> Path:
    repo = Path(tempfile.mkdtemp()) / "repo"
    repo.mkdir(parents=True)
    _r(["git", "init"], cwd=repo)
    _r(["git", "config", "user.email", "test@test.com"], cwd=repo)
    _r(["git", "config", "user.name", "Test"], cwd=repo)
    wf = repo / ".github" / "workflows"
    wf.mkdir(parents=True, exist_ok=True)
    for n, c in workflows.items():
        (wf / n).write_text(c, encoding="utf-8")
    _r(["git", "add", "."], cwd=repo, check=True)
    _r(["git", "commit", "-m", "base"], cwd=repo, check=True)
    return repo


def _gate(repo, threshold=24, check=False):
    cmd = [sys.executable, str(GATE), "--repo", str(repo), "--threshold", str(threshold)]
    if check:
        cmd.append("--check")
    return SUBP(cmd, capture_output=True, text=True, check=False)


class TestFclDetection(unittest.TestCase):
    """DetecciÃ³n de gates fail-closed."""

    def _wf(self, *names):
        return {"ci.yml": "jobs:\n" + "\n".join(f"  {n}:\n    runs-on: ubuntu-latest\n    steps:\n      - run: echo" for n in names)}

    def test_security_is_fcl(self):
        r = _mk_repo(self._wf("security"))
        p = _gate(r)
        self.assertEqual(p.returncode, 1, p.stdout + p.stderr)
        self.assertIn("security", p.stdout)

    def test_gate_named_is_fcl(self):
        r = _mk_repo(self._wf("deploy-gate", "lint"))
        p = _gate(r)
        self.assertEqual(p.returncode, 1, p.stdout + p.stderr)
        self.assertIn("deploy-gate", p.stdout)

    def test_no_fcl_jobs(self):
        r = _mk_repo(self._wf("lint", "build"))
        p = _gate(r)
        self.assertEqual(p.returncode, 2, p.stdout + p.stderr)
        self.assertIn("SKIP", p.stdout)

    def test_needs_of_deploy(self):
        # El job `security` está en needs de deploy: es FCL.
        wf = """jobs:
  security:
    runs-on: ubuntu-latest
    steps:
      - run: pip-audit
  deploy:
    needs: [security]
    runs-on: ubuntu-latest
    steps:
      - run: echo deploy
"""
        r = _mk_repo({"ci.yml": wf})
        p = _gate(r)
        self.assertEqual(p.returncode, 1, p.stdout + p.stderr)
        self.assertIn("security", p.stdout)

    def test_gate_kw(self):
        r = _mk_repo(self._wf("verify-e2e"))
        p = _gate(r)
        self.assertEqual(p.returncode, 1, p.stdout + p.stderr)
        self.assertIn("verify-e2e", p.stdout)

    def test_health_kw(self):
        r = _mk_repo(self._wf("health-check"))
        p = _gate(r)
        self.assertEqual(p.returncode, 1, p.stdout + p.stderr)
        self.assertIn("health-check", p.stdout)


class TestLivenessConfig(unittest.TestCase):
    """ConfiguraciÃ³n de liveness."""

    def test_no_liveness_is_finding(self):
        r = _mk_repo({"ci.yml": "jobs:\n  security:\n    runs-on: ubuntu-latest\n    steps:\n      - run: pip-audit\n"})
        p = _gate(r)
        self.assertEqual(p.returncode, 1, p.stdout + p.stderr)
        self.assertIn("FINDING", p.stdout)

    def test_liveness_asset_detected(self):
        r = _mk_repo({"ci.yml": "jobs:\n  security:\n    runs-on: ubuntu-latest\n    steps:\n      - run: pip-audit\n"})
        (r / "assets" / "gate-liveness").mkdir(parents=True)
        (r / "assets" / "gate-liveness" / "README.md").write_text("# config\n", encoding="utf-8")
        p = _gate(r)
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertIn("OK", p.stdout)

    def test_liveness_in_workflow(self):
        wf = """jobs:
  security:
    runs-on: ubuntu-latest
    steps:
      - run: pip-audit
  gate-liveness:
    runs-on: ubuntu-latest
    steps:
      - run: check-liveness
"""
        r = _mk_repo({"ci.yml": wf})
        p = _gate(r)
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertIn("OK", p.stdout)


class TestReviewDefects(unittest.TestCase):
    """Defectos hallados al cablear el gate en APAP (#4282): un comentario no es
    liveness, solo cuentan los jobs de primer nivel y un gate declarado en
    required_jobs tiene que invocarlo algun workflow."""

    SEC = "jobs:\n  security:\n    runs-on: ubuntu-latest\n    steps:\n      - run: pip-audit\n"

    def _with_liveness(self, r):
        (r / "assets" / "gate-liveness").mkdir(parents=True)
        (r / "assets" / "gate-liveness" / "README.md").write_text("# config\n", encoding="utf-8")

    def test_comment_mentioning_liveness_does_not_count(self):
        r = _mk_repo({"ci.yml": "# liveness: aviso tras 24h (pendiente)\n" + self.SEC})
        p = _gate(r)
        self.assertEqual(p.returncode, 1, p.stdout + p.stderr)
        self.assertIn("FINDING", p.stdout)

    def test_real_liveness_step_still_counts(self):
        wf = self.SEC.replace("      - run: pip-audit\n",
                              "      - name: liveness alert\n        run: pip-audit\n")
        r = _mk_repo({"ci.yml": wf})
        p = _gate(r)
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertIn("OK", p.stdout)

    def test_nested_keys_are_not_jobs(self):
        wf = ("jobs:\n  lint:\n    runs-on: ubuntu-latest\n    env:\n"
              "      deploy-token: x\n    steps:\n      - run: echo\n")
        r = _mk_repo({"ci.yml": wf})
        p = _gate(r)
        self.assertEqual(p.returncode, 2, p.stdout + p.stderr)
        self.assertNotIn("deploy-token", p.stdout)

    def test_declared_required_job_not_invoked_is_finding(self):
        r = _mk_repo({"ci.yml": self.SEC})
        self._with_liveness(r)
        (r / ".github" / "required-jobs-policy.json").write_text(
            json.dumps({"required_jobs": ["security", "deploy-gate"]}), encoding="utf-8")
        p = _gate(r)
        self.assertEqual(p.returncode, 1, p.stdout + p.stderr)
        self.assertIn("FINDING", p.stdout)
        self.assertIn("deploy-gate", p.stdout)

    def test_declared_required_job_invoked_passes(self):
        r = _mk_repo({"ci.yml": self.SEC})
        self._with_liveness(r)
        (r / ".github" / "required-jobs-policy.json").write_text(
            json.dumps({"required_jobs": ["security"]}), encoding="utf-8")
        p = _gate(r)
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)

    def test_unreadable_policy_is_invalid(self):
        r = _mk_repo({"ci.yml": self.SEC})
        self._with_liveness(r)
        (r / ".github" / "required-jobs-policy.json").write_text("{no json", encoding="utf-8")
        p = _gate(r)
        self.assertEqual(p.returncode, 2, p.stdout + p.stderr)
        self.assertIn("INVALID", p.stdout)


class TestThresholdConfig(unittest.TestCase):
    """Umbral configurado."""

    def test_default_is_24(self):
        r = _mk_repo({"ci.yml": "jobs:\n  security:\n    runs-on: ubuntu-latest\n    steps:\n      - run: pip-audit\n"})
        p = _gate(r)
        self.assertIn("24", p.stdout)

    def test_custom_threshold(self):
        r = _mk_repo({"ci.yml": "jobs:\n  security:\n    runs-on: ubuntu-latest\n    steps:\n      - run: pip-audit\n"})
        (r / "assets" / "gate-liveness").mkdir(parents=True)
        p = _gate(r, threshold=12)
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertIn("12", p.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
