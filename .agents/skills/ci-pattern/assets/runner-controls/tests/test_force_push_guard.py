#!/usr/bin/env python3
# ci-pattern asset tests — force-push / branch deletion guard (DysTelefonica/team-skills#303)
"""Force-push guard suite (black-box): ``forced`` or ``deleted`` on a
contract-protected branch opens an incident issue with the evidence
(exit 1); ordinary pushes and unprotected branches exit 0; a broken
environment fails closed (exit 2). The API is a local HTTP stub with the real
shapes; stdlib only, no network.

#1266: el merge gobernado despacha este guard por ``workflow_dispatch``
(#817), y ese evento no trae ``forced``/``deleted`` ni el rango de nivel
superior: el rango viaja en ``inputs``. El guard lee el rango con el lector
compartido ``push_range`` (#933) y dice en su salida que el evento no trae las
banderas, en vez de callar un veredicto que no midió.
"""
import argparse
import contextlib
import io
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from _github_stub import github_stub, stub_env
from test_push_compliance import BASE, REPO, SHA, contract, push_event

ASSET_DIR = Path(__file__).resolve().parent.parent
ASSET = ASSET_DIR / "check_force_push_guard.py"


def dispatch_event(before: str = BASE, after: str = SHA, ref: str = "refs/heads/main") -> dict:
    """Evento real de ``workflow_dispatch`` (#817): el rango viaja en ``inputs``.

    El payload de GitHub no trae ``forced``/``deleted`` (no existen en ese
    evento) ni ``before``/``after`` de nivel superior: el lanzador declara el
    rango auditado en ``on.workflow_dispatch.inputs`` y el workflow lo pasa con
    ``--event "$GITHUB_EVENT_PATH"``.
    """
    return {"event_name": "workflow_dispatch", "ref": ref,
            "inputs": {"before": before, "after": after},
            "workflow": "Force-push guard",
            "repository": {"full_name": REPO, "default_branch": "main"},
            "sender": {"login": "ardelperal"}}


def import_guard():
    """Importa el asset en proceso (el directorio del asset va al ``sys.path``)."""
    if str(ASSET_DIR) not in sys.path:
        sys.path.insert(0, str(ASSET_DIR))
    import check_force_push_guard
    return check_force_push_guard


def run_gate(root: Path, event: dict, contract_data: dict, state: dict) -> subprocess.CompletedProcess:
    (root / "event.json").write_text(json.dumps(event), encoding="utf-8")
    (root / "contract.json").write_text(json.dumps(contract_data), encoding="utf-8")
    with github_stub(root / "state.json", state) as base:
        return subprocess.run([sys.executable, str(ASSET), "--event", str(root / "event.json"),
                               "--contract", str(root / "contract.json")],
                              capture_output=True, text=True, env=stub_env(base), timeout=60)


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

    # #1266: el dispatch del merge gobernado es el evento que el guard recibe
    # en cada merge; antes de este fix fallaba cerrado con exit 2 y dejaba un
    # rojo permanente sin informacion en el commit de fusion.

    def test_dispatch_audits_the_range_and_exits_zero(self) -> None:
        proc = run_gate(self.root, dispatch_event(), contract(), guard_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn(BASE, proc.stdout)
        self.assertIn(SHA, proc.stdout)
        self.assertIn("no force-push nor deletion", proc.stdout)
        # La ausencia de las banderas se dice: un dispatch no las trae y eso no
        # es una medicion que haya pasado.
        self.assertIn("'forced'/'deleted'", proc.stdout)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_dispatch_with_malformed_range_fails_closed(self) -> None:
        proc = run_gate(self.root, dispatch_event(after="not-a-commit-id"), contract(), guard_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail-closed", proc.stdout + proc.stderr)
        self.assertIn("40-char commit ids", proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_push_keeps_the_same_verdict_as_before(self) -> None:
        """El camino del push no cambia: mismo veredicto y mismo texto que antes."""
        proc = run_gate(self.root, push_event(), contract(), guard_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("OK: push on protected branch 'main' is neither forced nor a deletion", proc.stdout)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_guard_reads_the_range_through_the_shared_reader(self) -> None:
        """Un solo lector del rango: el que ya usa el detective post-push (#933)."""
        guard = import_guard()
        import check_push_compliance
        self.assertIs(guard.push_range, check_push_compliance.push_range,
                      "el guard debe usar el lector compartido `push_range`")
        self.assertFalse(hasattr(guard, "SHA_RE"),
                         "un SHA_RE propio en el guard seria un segundo lector del rango")
        event = dispatch_event()
        (self.root / "event.json").write_text(json.dumps(event), encoding="utf-8")
        (self.root / "contract.json").write_text(json.dumps(contract()), encoding="utf-8")
        calls = []
        original = guard.push_range

        def recorder(payload: dict) -> tuple:
            calls.append(payload)
            return original(payload)

        guard.push_range = recorder
        try:
            with contextlib.redirect_stdout(io.StringIO()):
                code = guard.audit(argparse.Namespace(event=str(self.root / "event.json"),
                                                      contract=str(self.root / "contract.json")))
        finally:
            guard.push_range = original
        self.assertEqual(code, 0)
        self.assertEqual(calls, [event], "el guard debe leer el rango con `push_range`")

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
