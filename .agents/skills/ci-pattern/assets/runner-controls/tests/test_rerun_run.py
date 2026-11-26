#!/usr/bin/env python3
# ci-pattern asset tests -- relanzar un run caido con un solo comando (#1630)
"""``rerun_run.py`` relanza un run solo si sigue valiendo la pena.

Comprueba que el run es del repo, que termino en ``cancelled``, ``failure`` o
``timed_out`` y que su SHA sigue siendo la cabeza de su PR o de su rama; si no,
sale con 2 y dice el motivo. Si todo cuadra ejecuta ``gh run rerun <id> --failed``.
La ultima linea dice el resultado. El ``gh`` es siempre el falso (#1609).
"""
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tests"))
from _fake_gh import GH_ENV, PATH_TESTS_SKIP, install_fake_gh, run_guarded  # noqa: E402

ASSET = Path(__file__).resolve().parent.parent / "rerun_run.py"
REPO = "DysTelefonica/team-skills"
SHA = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
NEW = "9a8b7c6d5e4f3a2b1c0d9e8f7a6b5c4d3e2f1a0b"
RUN = 43488333183

STUB = '''
import json, sys
STATE = PLACEHOLDER
state = json.load(open(STATE, encoding="utf-8"))
args = sys.argv[1:]
state["calls"].append(args)
json.dump(state, open(STATE, "w", encoding="utf-8"))
if args[:2] == ["run", "rerun"]:
    sys.stderr.write(state.get("rerun_stderr", ""))
    sys.exit(state.get("rerun_rc", 0))
path = args[1] if args[:1] == ["api"] and len(args) > 1 else ""
if path in state["api"]:
    print(state["api"][path] if isinstance(state["api"][path], str) else json.dumps(state["api"][path]))
    sys.exit(0)
sys.stderr.write("gh: Not Found (HTTP 404)\\n")
sys.exit(1)
'''


def run_payload(**over: object) -> dict:
    payload: dict = {"id": RUN, "status": "completed", "conclusion": "cancelled",
                     "head_sha": SHA, "head_branch": "feat/x",
                     "repository": {"full_name": REPO}, "head_repository": {"full_name": REPO},
                     "pull_requests": [{"number": 1625}]}
    payload.update(over)
    return payload


def base_state(run: dict | None = None, pr_head: str = SHA) -> dict:
    return {"calls": [], "api": {
        f"repos/{REPO}/actions/runs/{RUN}": run or run_payload(),
        f"repos/{REPO}/pulls/1625": {"number": 1625, "state": "open", "head": {"sha": pr_head}},
        f"repos/{REPO}/git/ref/heads/feat/x": {"object": {"sha": SHA}},
    }}


class RerunRunTest(unittest.TestCase):
    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)

    def gate(self, state: dict, *extra: str, via_path: bool = False):
        state_file = self.root / "state.json"
        state_file.write_text(json.dumps(state), encoding="utf-8")
        shim = install_fake_gh(self.root / "bin", STUB.replace("PLACEHOLDER", repr(str(state_file))))
        env = {**os.environ, GH_ENV: str(shim)}
        if via_path:
            env.pop(GH_ENV)
            env["PATH"] = f"{shim.parent}{os.pathsep}{env.get('PATH', '')}"
        cmd = [sys.executable, str(ASSET), "--repo", REPO, "--run", str(RUN), *extra]
        return run_guarded(cmd, env=env, fixture_dir=self.root, capture_output=True,
                           text=True, encoding="utf-8", timeout=60)

    def reruns(self) -> list[list[str]]:
        calls = json.loads((self.root / "state.json").read_text(encoding="utf-8"))["calls"]
        return [c for c in calls if c[:2] == ["run", "rerun"]]

    def assert_last_line(self, proc, expected: str) -> None:
        self.assertEqual(proc.stdout.rstrip("\n").splitlines()[-1], expected, proc.stdout + proc.stderr)

    def test_ended_run_on_the_pr_head_is_rerun_with_failed(self) -> None:
        for conclusion in ("cancelled", "failure", "timed_out"):
            with self.subTest(conclusion=conclusion):
                proc = self.gate(base_state(run_payload(conclusion=conclusion)))
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
                self.assertEqual(self.reruns(), [["run", "rerun", str(RUN), "--failed", "-R", REPO]])
                self.assert_last_line(proc, f"RERUN: run {RUN} of {REPO} ({conclusion}, "
                                            f"{SHA[:12]}) relaunched with --failed")

    def test_all_jobs_reruns_the_whole_run(self) -> None:
        proc = self.gate(base_state(), "--all-jobs")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.reruns(), [["run", "rerun", str(RUN), "-R", REPO]])
        self.assert_last_line(proc, f"RERUN: run {RUN} of {REPO} (cancelled, {SHA[:12]}) "
                                    "relaunched with every job")

    def test_run_that_did_not_end_in_a_rerunnable_state_is_refused(self) -> None:
        cases = {"success": run_payload(conclusion="success"),
                 "in_progress": run_payload(status="in_progress", conclusion=None)}
        for label, payload in cases.items():
            with self.subTest(label=label):
                proc = self.gate(base_state(payload))
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertEqual(self.reruns(), [])
                status = "completed/success" if label == "success" else "in_progress/None"
                self.assert_last_line(proc, f"REFUSED: run {RUN} is {status}; only a completed run that "
                                            "ended cancelled, failure or timed_out is rerun")

    def test_run_of_another_repo_is_refused(self) -> None:
        proc = self.gate(base_state(run_payload(repository={"full_name": "other/repo"})))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.reruns(), [])
        self.assert_last_line(proc, f"REFUSED: run {RUN} belongs to other/repo, not to {REPO}")

    def test_unknown_run_fails_closed(self) -> None:
        state = base_state()
        del state["api"][f"repos/{REPO}/actions/runs/{RUN}"]
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.reruns(), [])
        self.assertTrue(proc.stdout.rstrip().splitlines()[-1].startswith(
            f"REFUSED: run {RUN} of {REPO} could not be read: gh api failed (1)"), proc.stdout)

    def test_run_whose_sha_is_no_longer_the_pr_head_says_to_push(self) -> None:
        proc = self.gate(base_state(pr_head=NEW))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.reruns(), [])
        self.assert_last_line(proc, f"REFUSED: run {RUN} ran on {SHA[:12]}, but the head of PR #1625 "
                                    f"is now {NEW[:12]}: push instead; a rerun repeats the old SHA")

    def test_run_without_pr_is_checked_against_its_branch(self) -> None:
        state = base_state(run_payload(pull_requests=[]))
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.reruns()), 1)
        state["api"][f"repos/{REPO}/git/ref/heads/feat/x"] = {"object": {"sha": NEW}}
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.reruns(), [])
        self.assert_last_line(proc, f"REFUSED: run {RUN} ran on {SHA[:12]}, but the head of branch "
                                    f"'feat/x' is now {NEW[:12]}: push instead; a rerun repeats the old SHA")

    def test_rerun_rejected_by_github_exits_1_with_its_reason(self) -> None:
        state = base_state()
        state["rerun_rc"] = 1
        state["rerun_stderr"] = "HTTP 403: Resource not accessible by integration\n"
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(len(self.reruns()), 1)
        self.assert_last_line(proc, f"FAILED: gh run rerun {RUN} exited 1: HTTP 403: Resource not "
                                    "accessible by integration")

    def test_run_payload_that_is_not_the_api_shape_fails_closed(self) -> None:
        state = base_state()
        state["api"][f"repos/{REPO}/actions/runs/{RUN}"] = "[]"
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.reruns(), [])
        self.assert_last_line(proc, f"REFUSED: run {RUN} of {REPO} is not the API run shape")
        proc = self.gate(base_state(run_payload(head_sha=None, pull_requests=[])))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.reruns(), [])
        self.assert_last_line(proc, f"REFUSED: run {RUN} has no head SHA")

    @unittest.skipIf(os.name == "nt", PATH_TESTS_SKIP)
    def test_without_explicit_gh_the_gh_on_path_is_used(self) -> None:
        proc = self.gate(base_state(), via_path=True)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.reruns()), 1)


if __name__ == "__main__":
    unittest.main()
