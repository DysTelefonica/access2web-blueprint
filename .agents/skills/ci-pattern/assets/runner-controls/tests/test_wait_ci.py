#!/usr/bin/env python3
# ci-pattern asset tests -- esperar a la CI de una PR (DysTelefonica/team-skills#1502)
"""`wait_ci.py` nunca devuelve exito con checks ausentes o pendientes.

Contra el stub HTTP de GitHub (sin `gh`): sin checks, pendientes, en rojo, en
verde, el agregador nombrado que no llega y la suite que aun esta arrancando
cuando ya hay un run verde (team-skills #1500)."""
import subprocess
import sys
import tempfile
import unittest
from datetime import UTC, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _github_stub import github_stub, stub_env  # noqa: E402

ASSET = Path(__file__).resolve().parent.parent / "wait_ci.py"
REPO = "DysTelefonica/team-skills"
SHA = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
FAST = ("--poll-seconds", "0.05")


def run(name: str, status: str = "completed", conclusion: str | None = "success") -> dict:
    return {"id": abs(hash(name)) % 10000, "name": name, "status": status, "conclusion": conclusion}


def suite(status: str) -> dict:
    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    return {"id": 1, "status": status, "app": {"slug": "github-actions"}, "created_at": now}


def state(**extra: object) -> dict:
    base = {"pr_details": {"7": {"number": 7, "state": "open", "head": {"sha": SHA}}},
            "check_runs": {SHA: []}, "calls": []}
    base.update(extra)
    return base


class WaitCiTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def wait(self, data: dict, *extra: str) -> subprocess.CompletedProcess:
        with github_stub(self.root / "state.json", data) as base:
            return subprocess.run([sys.executable, str(ASSET), "--repo", REPO, "--pr", "7", *FAST, *extra],
                                  capture_output=True, text=True, encoding="utf-8",
                                  env=stub_env(base), timeout=60)

    def test_no_checks_ever_is_not_green(self) -> None:
        proc = self.wait(state(), "--appear-seconds", "0.3")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(f"WAITING: no check runs yet for {SHA[:12]}", proc.stdout)
        self.assertIn(f"NOT STARTED: no check runs appeared for {SHA[:12]} within 0.3 s", proc.stdout)
        self.assertNotIn("GREEN", proc.stdout)

    def test_waits_through_absent_and_pending_then_green(self) -> None:
        data = state(check_runs_script={SHA: [[], [run("unit", "in_progress", None)],
                                              [run("unit"), run("lint", conclusion="skipped")]]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("WAITING: 1 check run(s) pending: unit", proc.stdout)
        self.assertIn(f"GREEN: 2 check run(s) of {SHA[:12]} passed", proc.stdout)

    def test_a_red_run_is_exit_1_naming_it(self) -> None:
        data = state(check_runs={SHA: [run("unit"), run("e2e", conclusion="failure"),
                                       run("lint", conclusion="cancelled")]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("RED: e2e (failure), lint (cancelled)", proc.stdout)

    def test_named_aggregator_that_never_reports_times_out(self) -> None:
        data = state(check_runs={SHA: [run("unit")]})
        proc = self.wait(data, "--check", "required", "--timeout-seconds", "0.3")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("WAITING: required check(s) not reported yet: required", proc.stdout)
        self.assertIn("TIMEOUT: required check(s) not reported yet: required after 0.3 s", proc.stdout)
        self.assertNotIn("GREEN", proc.stdout)

    def test_a_green_run_with_a_suite_still_starting_is_not_green_yet(self) -> None:
        # team-skills #1500: pr-size ya en verde y la suite de `tests` en cola sin runs.
        data = state(check_runs_script={SHA: [[run("pr-size")], [run("pr-size"), run("unit")]]},
                     check_suites_script={SHA: [[suite("completed"), suite("queued")],
                                                [suite("completed"), suite("completed")]]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn(f"WAITING: 1 GitHub Actions check suite(s) still starting for {SHA[:12]}", proc.stdout)
        self.assertIn(f"GREEN: 2 check run(s) of {SHA[:12]} passed", proc.stdout)

    def test_pending_forever_times_out(self) -> None:
        data = state(check_runs={SHA: [run("unit", "queued", None)]})
        proc = self.wait(data, "--timeout-seconds", "0.3")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("TIMEOUT: 1 check run(s) pending: unit after 0.3 s", proc.stdout)

    def test_api_error_fails_closed(self) -> None:
        proc = self.wait(state(api_error=True))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_bad_arguments_are_a_usage_error(self) -> None:
        proc = subprocess.run([sys.executable, str(ASSET), "--repo", REPO, "--pr", "0"],
                              capture_output=True, text=True, timeout=30)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("--pr must be positive", proc.stderr)


if __name__ == "__main__":
    unittest.main()
