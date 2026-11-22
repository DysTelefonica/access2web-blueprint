#!/usr/bin/env python3
# ci-pattern asset tests -- el merge gobernado lee `required` del run del evento pull_request (DysTelefonica/team-skills#1713)
"""El veredicto de `required` sale de la ejecución del evento `pull_request` sobre el SHA de cabeza.

Un `required` de otro evento (un `workflow_dispatch` posterior sobre el mismo SHA) no cuenta:
ni para bloquear por un `skipped`, ni para desbloquear por un verde. Sin ningún run de
`pull_request` para el requerido, sigue faltando (falla cerrado). La lectura de las
ejecuciones de workflow falla cerrada (exit 2).
"""

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _pr_suite_stub import suites_github_stub  # noqa: E402
from test_governed_merge import base_state, contract, pull  # noqa: E402
from test_push_compliance import CHECK, REPO, SHA  # noqa: E402

ASSET = Path(__file__).resolve().parents[1] / "check_governed_merge.py"
PR_SUITE = 100  # check suite del run del evento pull_request
PRT_SUITE = 150  # check suite de un run pull_request_target (pr-gate, pr-contract; #1773)
DISPATCH_SUITE = 200  # check suite de un workflow_dispatch posterior, mismo SHA
TYPE_BUG = [{"name": "type:bug"}]


def check_run(run_id, conclusion, suite):
    return {"id": run_id, "name": CHECK, "status": "completed", "conclusion": conclusion,
            "html_url": f"https://github.com/{REPO}/actions/runs/{run_id}",
            "check_suite": {"id": suite}}


def workflow_run(run_id, event, suite):
    return {"id": run_id, "event": event, "head_sha": SHA, "check_suite_id": suite,
            "status": "completed"}


class TestRequiredFromThePullRequestRun(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)

    def gate(self, state, declared=None):
        contract_data = contract() if declared is None else declared
        (self.root / "contract.json").write_text(json.dumps(contract_data), encoding="utf-8")
        with suites_github_stub(self.root / "state.json", state) as base:
            from _github_stub import stub_env
            env = stub_env(base)
            env.pop("GITHUB_STEP_SUMMARY", None)
            return subprocess.run([sys.executable, str(ASSET), "--repo", REPO, "--pr", "7",
                                   "--contract", str(self.root / "contract.json"),
                                   "--max-lines", "400", "--wait-seconds", "0"],
                                  capture_output=True, text=True, env=env, timeout=60)

    def merges(self):
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))["merges"]

    def state(self, runs, workflow_runs):
        state = base_state()
        state["pr_details"] = {"7": pull(labels=TYPE_BUG)}
        state["check_runs"] = {SHA: runs}
        state["workflow_runs_for_sha"] = workflow_runs
        return state

    def test_a_skipped_required_from_a_dispatch_does_not_block_a_green_pull_request_run(self):
        # Acceptance literal (#1713): 'a skipped required from a dispatch on the same SHA
        # does not block a PR whose pull_request run is green'.
        state = self.state(
            runs=[check_run(5, "success", PR_SUITE), check_run(9, "skipped", DISPATCH_SUITE)],
            workflow_runs=[workflow_run(50, "pull_request", PR_SUITE),
                           workflow_run(90, "workflow_dispatch", DISPATCH_SUITE)])
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.merges()), 1)

    def test_a_red_required_from_the_pull_request_run_blocks_even_if_a_later_dispatch_is_green(self):
        state = self.state(
            runs=[check_run(5, "failure", PR_SUITE), check_run(9, "success", DISPATCH_SUITE)],
            workflow_runs=[workflow_run(50, "pull_request", PR_SUITE),
                           workflow_run(90, "workflow_dispatch", DISPATCH_SUITE)])
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])
        self.assertIn(f'required check "{CHECK}": failure', proc.stdout)

    def test_a_required_with_no_pull_request_run_stays_missing_and_blocks(self):
        # Un verde de dispatch no sustituye al run del pull_request: falta, y bloquea.
        state = self.state(
            runs=[check_run(9, "success", DISPATCH_SUITE)],
            workflow_runs=[workflow_run(90, "workflow_dispatch", DISPATCH_SUITE)])
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])
        self.assertIn("missing", proc.stdout)

    def test_a_green_required_from_a_pull_request_target_run_is_not_missing(self):
        # #1773: `pr-gate` y `pr-contract` los publica un workflow pull_request_target; su
        # check suite es de la PR y cuenta. Antes salia 'missing' con el check en verde.
        state = self.state(
            runs=[check_run(5, "success", PRT_SUITE)],
            workflow_runs=[workflow_run(50, "pull_request_target", PRT_SUITE)])
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.merges()), 1)
        self.assertNotIn("missing", proc.stdout)

    def test_a_red_required_from_a_pull_request_target_run_blocks(self):
        state = self.state(
            runs=[check_run(5, "failure", PRT_SUITE), check_run(9, "success", DISPATCH_SUITE)],
            workflow_runs=[workflow_run(50, "pull_request_target", PRT_SUITE),
                           workflow_run(90, "workflow_dispatch", DISPATCH_SUITE)])
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])
        self.assertIn(f'required check "{CHECK}": failure', proc.stdout)

    def test_an_unreadable_workflow_run_listing_fails_closed(self):
        state = self.state(
            runs=[check_run(5, "success", PR_SUITE)],
            workflow_runs=[workflow_run(50, "pull_request", PR_SUITE)])
        state["workflow_runs_error"] = True
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])

    def run_(self, name, conclusion):
        return {"id": abs(hash(name)) % 999 + 1, "name": name, "status": "completed",
                "conclusion": conclusion, "check_suite": {"id": PR_SUITE}}

    def advisory_state(self, runs):
        state = self.state(runs, [workflow_run(50, "pull_request", PR_SUITE)])
        return state

    def test_un_requerido_nunca_queda_tapado_por_un_patron(self):
        # #1762: el patron casa con el nombre del check, pero el check es requerido.
        declared = contract()
        declared["advisory_checks"] = ["*"]
        proc = self.gate(self.advisory_state([self.run_(CHECK, "failure")]), declared)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])

    def test_un_espejo_advisory_no_bloquea_el_merge(self):
        # #1762: un espejo con startup_failure no bloquea el merge gobernado.
        declared = contract()
        declared["advisory_checks"] = ["*-gh"]
        state = self.advisory_state([self.run_(CHECK, "success"),
                                      self.run_("docs-gh", "startup_failure")])
        proc = self.gate(state, declared)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.merges()), 1)

    def test_una_declaracion_invalida_falla_cerrado(self):
        for value in ("*-gh", [3], [""], None):
            with self.subTest(advisory_checks=value):
                declared = contract()
                declared["advisory_checks"] = value
                proc = self.gate(self.advisory_state([self.run_(CHECK, "success")]), declared)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertEqual(self.merges(), [])


if __name__ == "__main__":
    unittest.main()
