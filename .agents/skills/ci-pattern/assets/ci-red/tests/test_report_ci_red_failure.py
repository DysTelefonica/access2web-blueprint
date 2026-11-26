#!/usr/bin/env python3
# ci-pattern asset tests — ci-red: el rojo programado tiene dueño (#1508, HR-89)
"""Suite del aviso ``ci-red`` de los runs programados (HR-89): camino del rojo.

Cubre la issue que abre o comenta un run programado en rojo, el run fuera de
alcance y los errores de uso o de contrato. El camino del verde (comentar y
cerrar) vive en ``test_report_ci_red_close.py``. El ``gh`` simulado está en
``_gh_stub.py``.

Códigos de salida: 0 aviso dejado o nada que hacer · 1 ``gh`` no pudo dejar
el aviso · 2 uso o contrato del host inservible.
"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _gh_stub import CONTRACT, HostTestCase, open_issue  # noqa: E402


class ScheduledFailureTests(HostTestCase):
    def test_first_failure_opens_one_ci_red_issue_assigned_to_the_coordinator(self):
        host = self.host()
        proc = host.run("failure")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        issues = host.state["issues"]
        self.assertEqual(len(issues), 1, issues)
        issue = issues[0]
        self.assertEqual(issue["title"], "[ci-red] main-audit")
        self.assertEqual(issue["labels"], ["ci-red"])
        self.assertEqual(issue["assignees"], ["coordinador"])
        self.assertIn("https://github.com/example/consumer/actions/runs/4242", issue["body"])
        self.assertIn("0123456789abcdef0123456789abcdef01234567", issue["body"])
        self.assertIn("ci-red: issue #1 abierta para el workflow «main-audit»", proc.stdout)

    def test_second_failure_comments_on_the_open_issue_and_opens_no_other(self):
        host = self.host()
        self.assertEqual(host.run("failure").returncode, 0)
        proc = host.run("failure", GITHUB_RUN_ID="4243")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        issues = host.state["issues"]
        self.assertEqual(len(issues), 1, issues)
        self.assertEqual(len(issues[0]["comments"]), 1)
        self.assertIn("actions/runs/4243", issues[0]["comments"][0])
        self.assertIn("ci-red: comentario en la issue #1", proc.stdout)

    def test_the_issue_is_per_workflow(self):
        host = self.host(issues=[open_issue(7, "[ci-red] r2-smoke")])
        proc = host.run("failure")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        titles = sorted(i["title"] for i in host.state["issues"])
        self.assertEqual(titles, ["[ci-red] main-audit", "[ci-red] r2-smoke"])
        self.assertEqual(host.state["issues"][0]["comments"], [])

    def test_an_issue_without_the_ci_red_label_is_not_the_tracking_issue(self):
        host = self.host(issues=[open_issue(3, "[ci-red] main-audit", labels=("type:bug",))])
        self.assertEqual(host.run("failure").returncode, 0)
        self.assertEqual(len(host.state["issues"]), 2)

    def test_gh_failure_is_exit_1_never_a_silent_zero(self):
        host = self.host(fail=["issue", "create"])
        proc = host.run("failure")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("gh issue create", proc.stderr)
        self.assertEqual(host.state["issues"], [])


class OutOfScopeRunTests(HostTestCase):
    def test_other_branch_or_other_event_opens_nothing_and_calls_no_gh(self):
        for overrides in ({"GITHUB_REF_NAME": "feat/x"},
                          {"GITHUB_EVENT_NAME": "push"},
                          {"GITHUB_EVENT_NAME": "pull_request"},
                          {"GITHUB_EVENT_NAME": "workflow_dispatch"}):
            for outcome in ("failure", "success"):
                with self.subTest(overrides=overrides, outcome=outcome):
                    host = self.host(issues=[open_issue(5, "[ci-red] main-audit")])
                    proc = host.run(outcome, **overrides)
                    self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
                    self.assertIn("ci-red: nada que hacer", proc.stdout)
                    self.assertEqual(host.state["calls"], [])


class FailClosedTests(HostTestCase):
    def assert_exit_2(self, host, needle, **env_overrides):
        proc = host.run("failure", **env_overrides)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn(needle, proc.stderr)
        self.assertEqual(host.state["calls"], [])

    def test_contract_without_coordinator_is_exit_2(self):
        contract = dict(CONTRACT)
        del contract["coordinator"]
        self.assert_exit_2(self.host(contract=contract), "coordinator")

    def test_contract_with_an_empty_or_non_text_coordinator_is_exit_2(self):
        for bad in ("", 7, ["a"]):
            with self.subTest(coordinator=bad):
                self.assert_exit_2(self.host(contract={**CONTRACT, "coordinator": bad}),
                                   "coordinator")

    def test_contract_without_the_ci_red_label_is_exit_2(self):
        contract = {**CONTRACT, "labels": [{"name": "chain:partial", "class": "host-enforced"}]}
        self.assert_exit_2(self.host(contract=contract), "ci-red")

    def test_unreadable_contract_is_exit_2(self):
        self.assert_exit_2(self.host(contract="{no json"), "host-contract")

    def test_missing_run_identity_is_exit_2(self):
        self.assert_exit_2(self.host(), "GITHUB_WORKFLOW", GITHUB_WORKFLOW="")

    def test_unknown_outcome_is_a_usage_error(self):
        host = self.host()
        proc = host.run("cancelled")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(host.state["calls"], [])


if __name__ == "__main__":
    unittest.main()
