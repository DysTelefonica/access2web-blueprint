#!/usr/bin/env python3
# ci-pattern asset tests — ci-red: el verde cierra la issue (#1508, HR-89)
"""Suite del aviso ``ci-red`` de los runs programados (HR-89): camino del verde.

Un run programado en verde comenta y cierra la issue ``ci-red`` abierta de su
workflow, y sin issue abierta no cambia nada. El ``gh`` simulado está en
``_gh_stub.py``.
"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _gh_stub import HostTestCase, open_issue  # noqa: E402


class ScheduledSuccessTests(HostTestCase):
    def test_green_run_comments_and_closes_the_open_issue_of_its_workflow(self):
        host = self.host(issues=[open_issue(5, "[ci-red] main-audit"),
                                 open_issue(6, "[ci-red] r2-smoke")])
        proc = host.run("success")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        by_number = {i["number"]: i for i in host.state["issues"]}
        self.assertEqual(by_number[5]["state"], "closed")
        self.assertEqual(len(by_number[5]["comments"]), 1)
        self.assertIn("actions/runs/4242", by_number[5]["comments"][0])
        self.assertEqual(by_number[6]["state"], "open")
        self.assertIn("ci-red: issue #5 cerrada", proc.stdout)

    def test_green_run_without_open_issue_changes_nothing(self):
        host = self.host()
        proc = host.run("success")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(host.state["issues"], [])
        self.assertEqual([c[:2] for c in host.state["calls"]], [["issue", "list"]])


if __name__ == "__main__":
    unittest.main()
