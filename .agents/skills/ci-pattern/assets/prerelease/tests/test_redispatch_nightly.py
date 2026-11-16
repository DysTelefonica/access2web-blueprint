#!/usr/bin/env python3
# ci-pattern asset tests -- relanzar la canalizacion nocturna sobre main (DysTelefonica/team-skills#1627, slice 3)
"""``redispatch_nightly.py``: relanza el workflow nocturno; ``--dry-run`` no llama a gh.

Es idempotente: si ya hay una ejecucion nocturna activa para la cabeza de main, no
relanza (SKIP). El ``gh`` es siempre el falso (#1609); el caso «gh no existe» usa una
ruta inexistente y nunca ejecuta un gh real.
"""
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tests"))
from _fake_gh import GH_ENV, install_fake_gh, run_guarded  # noqa: E402

ASSETS = Path(__file__).resolve().parents[1]
ASSET = ASSETS / "redispatch_nightly.py"
sys.path.insert(0, str(ASSETS))
import redispatch_nightly as rn  # noqa: E402

REPO = "DysTelefonica/team-skills"
MAIN = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
EXPECTED_ARGS = ["workflow", "run", "nightly-pipeline.yml", "--repo", REPO, "--ref", "main"]

# Estado: main_sha, runs_before (ejecuciones ya activas), rc/stderr del dispatch, fail_reads.
# Tras un `workflow run` correcto aparece una ejecucion en cola, como hace GitHub.
STUB = '''
import json, sys
STATE = PLACEHOLDER
state = json.load(open(STATE, encoding="utf-8"))
args = sys.argv[1:]
state["calls"].append(args)
if args[:2] == ["workflow", "run"]:
    if state.get("rc", 0) == 0:
        state["dispatched"] = state.get("dispatched", 0) + 1
    json.dump(state, open(STATE, "w", encoding="utf-8"))
    sys.stderr.write(state.get("stderr", ""))
    sys.exit(state.get("rc", 0))
json.dump(state, open(STATE, "w", encoding="utf-8"))
path = args[1] if args[:1] == ["api"] and len(args) > 1 else ""
if state.get("fail_reads"):
    sys.stderr.write("gh: Service Unavailable (HTTP 503)\\n")
    sys.exit(1)
if path.endswith("/git/ref/heads/main"):
    print(json.dumps({"object": {"sha": state["main_sha"]}}))
    sys.exit(0)
if "/actions/workflows/nightly-pipeline.yml/runs" in path:
    runs = list(state.get("runs_before", []))
    if state.get("dispatched"):
        runs.append({"id": 77, "status": "queued", "head_sha": state["main_sha"]})
    print(json.dumps({"workflow_runs": runs}))
    sys.exit(0)
sys.stderr.write("gh: Not Found (HTTP 404)\\n")
sys.exit(1)
'''


class RedispatchCliTest(unittest.TestCase):
    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.state = {"calls": [], "main_sha": MAIN}

    def run_times(self, extra: list[str], times: int = 1):
        """Corre el CLI ``times`` veces sobre el mismo estado; devuelve (procesos, llamadas)."""
        state_file = self.root / "state.json"
        state_file.write_text(json.dumps(self.state), encoding="utf-8")
        shim = install_fake_gh(self.root / "bin", STUB.replace("PLACEHOLDER", repr(str(state_file))))
        env = {**os.environ, GH_ENV: str(shim)}
        procs = [run_guarded([sys.executable, str(ASSET), "--repo", REPO, *extra], env=env,
                             fixture_dir=self.root, capture_output=True, text=True, encoding="utf-8", timeout=60)
                 for _ in range(times)]
        calls = json.loads(state_file.read_text(encoding="utf-8"))["calls"]
        return procs, calls

    @staticmethod
    def dispatches(calls: list) -> list:
        return [c for c in calls if c[:2] == ["workflow", "run"]]

    def test_dry_run_twice_gives_identical_output_and_no_gh_call(self) -> None:
        (first, second), calls = self.run_times(["--dry-run"], times=2)
        self.assertEqual(first.stdout, second.stdout)
        self.assertEqual((first.returncode, second.returncode), (0, 0))
        self.assertEqual(calls, [], "dry-run must not invoke gh at all")
        self.assertTrue(first.stdout.splitlines()[-1].startswith("DRY-RUN: would dispatch"), first.stdout)

    def test_real_run_dispatches_the_workflow_on_main(self) -> None:
        (proc,), calls = self.run_times([])
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.dispatches(calls), [EXPECTED_ARGS])
        self.assertTrue(proc.stdout.splitlines()[-1].startswith("REDISPATCHED: "), proc.stdout)

    def test_real_run_twice_dispatches_exactly_once(self) -> None:
        (first, second), calls = self.run_times([], times=2)
        self.assertEqual((first.returncode, second.returncode), (0, 0), first.stdout + second.stdout)
        self.assertEqual(len(self.dispatches(calls)), 1, calls)
        self.assertTrue(second.stdout.splitlines()[-1].startswith("SKIP: "), second.stdout)
        self.assertIn("77", second.stdout)

    def test_active_nightly_for_the_same_sha_is_not_dispatched_again(self) -> None:
        for status in ("queued", "in_progress", "waiting", "action_required"):
            with self.subTest(status=status):
                self.state = {"calls": [], "main_sha": MAIN, "runs_before": [{"id": 5, "status": status}]}
                (proc,), calls = self.run_times([])
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
                self.assertEqual(self.dispatches(calls), [])
                self.assertTrue(proc.stdout.splitlines()[-1].startswith("SKIP: "), proc.stdout)

    def test_finished_nightly_does_not_block_a_new_dispatch(self) -> None:
        self.state = {"calls": [], "main_sha": MAIN, "runs_before": [{"id": 5, "status": "completed"}]}
        (proc,), calls = self.run_times([])
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.dispatches(calls), [EXPECTED_ARGS])

    def test_unreadable_state_refuses_without_dispatching(self) -> None:
        self.state = {"calls": [], "main_sha": MAIN, "fail_reads": True}
        (proc,), calls = self.run_times([])
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.dispatches(calls), [])
        self.assertTrue(proc.stdout.splitlines()[-1].startswith("REFUSED: "), proc.stdout)

    def test_malformed_repo_is_refused_without_gh_calls(self) -> None:
        for repo in ("", "justname", "a/b/c", "owner/ /x", "o/r\n", "../..", "owner/.", "./repo"):
            with self.subTest(repo=repo):
                state_file = self.root / "state.json"
                state_file.write_text(json.dumps(self.state), encoding="utf-8")
                shim = install_fake_gh(self.root / "bin", STUB.replace("PLACEHOLDER", repr(str(state_file))))
                env = {**os.environ, GH_ENV: str(shim)}
                proc = run_guarded([sys.executable, str(ASSET), "--repo", repo], env=env,
                                   fixture_dir=self.root, capture_output=True, text=True,
                                   encoding="utf-8", timeout=60)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertTrue(proc.stdout.splitlines()[-1].startswith("REFUSED: "), proc.stdout)
                self.assertEqual(json.loads(state_file.read_text(encoding="utf-8"))["calls"], [])

    def test_github_rejection_exits_1_with_its_reason(self) -> None:
        self.state = {"calls": [], "main_sha": MAIN, "rc": 1,
                      "stderr": "HTTP 403: Resource not accessible by integration\n"}
        (proc,), calls = self.run_times([])
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.dispatches(calls), [EXPECTED_ARGS])
        self.assertTrue(proc.stdout.splitlines()[-1].startswith("FAILED: "), proc.stdout)
        self.assertIn("HTTP 403: Resource not accessible by integration", proc.stdout.splitlines()[-1])

    def test_gh_that_cannot_run_refuses(self) -> None:
        # Ruta inexistente: el proceso no arranca y no se ejecuta ningun gh real.
        old = os.environ.get(GH_ENV)
        os.environ[GH_ENV] = str(self.root / "no-such-gh")
        try:
            code, line = rn.redispatch(REPO, dry_run=False)
        finally:
            if old is None:
                os.environ.pop(GH_ENV, None)
            else:
                os.environ[GH_ENV] = old
        self.assertEqual(code, 2)
        self.assertTrue(line.startswith("REFUSED: "), line)


if __name__ == "__main__":
    unittest.main()
