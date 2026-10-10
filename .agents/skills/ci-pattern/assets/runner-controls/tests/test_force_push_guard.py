#!/usr/bin/env python3
# ci-pattern asset tests — force-push / branch deletion guard (DysTelefonica/team-skills#303)
"""Force-push guard suite (black-box): ``forced`` or ``deleted`` on a
contract-protected branch opens an incident issue with the evidence
(exit 1); ordinary pushes and unprotected branches exit 0; a broken
environment fails closed (exit 2). ``gh`` simulated with the real API
shapes; stdlib only, no network."""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from test_push_compliance import BASE, GH_SHIM, REPO, SHA, contract, push_event

ASSET = Path(__file__).resolve().parent.parent / "check_force_push_guard.py"


def run_gate(root: Path, event: dict, contract_data: dict, state: dict) -> subprocess.CompletedProcess:
    (root / "event.json").write_text(json.dumps(event), encoding="utf-8")
    (root / "contract.json").write_text(json.dumps(contract_data), encoding="utf-8")
    (root / "state.json").write_text(json.dumps(state), encoding="utf-8")
    shim = root / "bin" / "gh"
    shim.parent.mkdir(exist_ok=True)
    shim.write_text(GH_SHIM.format(state_path=str(root / "state.json")), encoding="utf-8")
    shim.chmod(0o755)
    env = {**os.environ, "PATH": f"{root / 'bin'}{os.pathsep}{os.environ.get('PATH', '')}"}
    return subprocess.run([sys.executable, str(ASSET), "--event", str(root / "event.json"),
                           "--contract", str(root / "contract.json")],
                          capture_output=True, text=True, env=env, timeout=60)


def guard_state() -> dict:
    return {"created_issues": [], "calls": [], "merges": []}


class ForcePushGuardTests(unittest.TestCase):
    """#303 control 3: force-push o borrado de rama protegida abre incidente."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def created_issue(self) -> dict:
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(len(state["created_issues"]), 1, state)
        return state["created_issues"][0]

    def test_ordinary_push_exits_zero_without_issue(self) -> None:
        proc = run_gate(self.root, push_event(), contract(), guard_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_force_push_to_protected_branch_opens_incident(self) -> None:
        event = push_event()
        event["forced"] = True
        proc = run_gate(self.root, event, contract(), guard_state())
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        issue = self.created_issue()
        self.assertIn(BASE, issue["body"])
        self.assertIn(SHA, issue["body"])
        self.assertIn("force", (issue["title"] + issue["body"]).lower())
        self.assertIn("INCIDENT ISSUE", proc.stdout + proc.stderr)

    def test_protected_branch_deletion_opens_incident(self) -> None:
        event = push_event()
        event["deleted"] = True
        event["after"] = "0" * 40
        proc = run_gate(self.root, event, contract(), guard_state())
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        issue = self.created_issue()
        self.assertIn(BASE, issue["body"])
        self.assertIn("deleted", issue["title"] + issue["body"])

    def test_force_push_to_unprotected_branch_exits_zero(self) -> None:
        event = push_event(ref="refs/heads/topic/x")
        event["forced"] = True
        proc = run_gate(self.root, event, contract(), guard_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_fail_closed_environment(self) -> None:
        event = push_event()
        event["forced"] = True
        del event["repository"]
        proc = run_gate(self.root, event, contract(), guard_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        state = guard_state()
        state["issue_post_error"] = True
        proc = run_gate(self.root, {**push_event(), "forced": True}, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail", (proc.stdout + proc.stderr).lower())


if __name__ == "__main__":
    unittest.main()
