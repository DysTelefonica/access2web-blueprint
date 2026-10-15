#!/usr/bin/env python3
# ci-pattern asset tests — gate de perfil de runners (team-skills#435).
"""check_runner_profile.py: en perfil `self-hosted`, todo job de un
workflow disparado por pull_request/push corre en el pool declarado
(P52) o figura en P53 con motivo. Sin eso, rojo.

Criterios de la issue (S5): job hosted no permitido → rojo; job en el
pool → verde; excepción justificada → verde. Fail-closed (HR-3):
ci-pattern.yaml ausente/ilegible, P51 inválido o perfil self-hosted sin
P52 → exit 2. Perfil `hosted` (default) → gate inocuo, exit 0.

Rodeado con repos mínimos en tmp; nunca contra este clon.
"""
from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_runner_profile.py"


def run_gate(repo: Path):
    return subprocess.run(
        [sys.executable, str(GATE), str(repo)],
        capture_output=True, text=True, timeout=60,
    )


class Repo:
    """Repo mínimo: ci-pattern.yaml + workflows bajo demanda."""

    def __init__(self, root: Path):
        self.root = root
        self.wf_dir = root / ".github" / "workflows"

    def params(self, text: str):
        (self.root / "ci-pattern.yaml").write_text(text, encoding="utf-8")

    def workflow(self, name: str, body: str):
        self.wf_dir.mkdir(parents=True, exist_ok=True)
        (self.wf_dir / name).write_text(body, encoding="utf-8")


SELF_HOSTED_PARAMS = """\
P51_runners_profile: self-hosted
P52_runners_pool_labels: [self-hosted, Linux, apap-ci]
"""

POOL_JOB = """\
on:
  pull_request:
jobs:
  lint:
    runs-on: [self-hosted, Linux, apap-ci]
    steps:
      - run: echo ok
"""

HOSTED_JOB = """\
on: [pull_request, push]
jobs:
  lint:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
"""


class RunnerProfileGateTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="rprof-"))
        self.repo = Repo(self.tmp)

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)

    # ---- los tres casos del criterio de aceptacion S5 ----------------

    def test_hosted_job_not_allowed_is_red(self):
        self.repo.params(SELF_HOSTED_PARAMS)
        self.repo.workflow("ci.yml", HOSTED_JOB)
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 1, f"stdout={r.stdout} stderr={r.stderr}")
        self.assertIn("HOSTED-SIN-EXCEPCION", r.stdout)
        self.assertIn("ci.yml", r.stdout)

    def test_pool_job_is_green(self):
        self.repo.params(SELF_HOSTED_PARAMS)
        self.repo.workflow("ci.yml", POOL_JOB)
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 0, f"stdout={r.stdout} stderr={r.stderr}")

    def test_justified_exception_is_green(self):
        self.repo.params(
            SELF_HOSTED_PARAMS
            + "P53_hosted_exceptions:\n  lint: el build nativo de x86 no existe en el pool\n"
        )
        self.repo.workflow("ci.yml", HOSTED_JOB)
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 0, f"stdout={r.stdout} stderr={r.stderr}")

    # ---- relajacion sin motivo (estilo HR-15) -----------------------

    def test_exception_without_reason_is_red(self):
        self.repo.params(
            SELF_HOSTED_PARAMS + "P53_hosted_exceptions:\n  lint:\n"
        )
        self.repo.workflow("ci.yml", HOSTED_JOB)
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 1, f"stdout={r.stdout}")
        self.assertIn("MOTIVO-VACIO", r.stdout)

    # ---- pool declarado vs fuera de pool -----------------------------

    def test_self_hosted_outside_declared_pool_is_red(self):
        self.repo.params(SELF_HOSTED_PARAMS)
        self.repo.workflow(
            "ci.yml",
            "on: push\njobs:\n  x:\n    runs-on: [self-hosted, other-pool]\n",
        )
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 1, f"stdout={r.stdout}")
        self.assertIn("FUERA-DE-POOL", r.stdout)

    def test_job_without_runs_on_is_red(self):
        # Sin runs-on heredaria el default hosted (ubuntu-latest).
        self.repo.params(SELF_HOSTED_PARAMS)
        self.repo.workflow("ci.yml", "on: pull_request\njobs:\n  x:\n    steps:\n      - run: echo ok\n")
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 1, f"stdout={r.stdout}")
        self.assertIn("SIN-RUNS-ON", r.stdout)

    def test_dynamic_runs_on_doubt_is_red(self):
        self.repo.params(SELF_HOSTED_PARAMS)
        self.repo.workflow(
            "ci.yml",
            "on: pull_request\njobs:\n  x:\n    runs-on: ${{ matrix.os }}\n",
        )
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 1, f"stdout={r.stdout}")
        self.assertIn("RUNS-ON-DINAMICO", r.stdout)

    # ---- ambito: solo workflows con pull_request/push ----------------

    def test_schedule_only_workflow_not_gated(self):
        self.repo.params(SELF_HOSTED_PARAMS)
        self.repo.workflow("nightly.yml", "on:\n  schedule:\n    - cron: '0 4 * * *'\njobs:\n  x:\n    runs-on: ubuntu-24.04\n")
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 0, f"stdout={r.stdout}")

    def test_workflow_level_defaults_pool_is_green(self):
        self.repo.params(SELF_HOSTED_PARAMS)
        self.repo.workflow(
            "ci.yml",
            "on: pull_request\ndefaults:\n  jobs:\n    runs-on: [self-hosted, Linux, apap-ci]\njobs:\n  x:\n    steps:\n      - run: echo ok\n",
        )
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 0, f"stdout={r.stdout}")

    # ---- default inocuo y fail-closed --------------------------------

    def test_hosted_profile_makes_gate_inert(self):
        self.repo.params("P01_branch_name_pattern: x\n")  # sin P51: default hosted
        self.repo.workflow("ci.yml", HOSTED_JOB)
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 0, f"stdout={r.stdout}")

    def test_missing_params_file_is_invalid(self):
        self.repo.workflow("ci.yml", HOSTED_JOB)
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 2, f"stdout={r.stdout} stderr={r.stderr}")

    def test_unknown_profile_value_is_invalid(self):
        self.repo.params("P51_runners_profile: nube\n")
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 2)

    def test_self_hosted_profile_without_pool_is_invalid(self):
        self.repo.params("P51_runners_profile: self-hosted\n")
        self.repo.workflow("ci.yml", POOL_JOB)
        r = run_gate(self.tmp)
        self.assertEqual(r.returncode, 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
