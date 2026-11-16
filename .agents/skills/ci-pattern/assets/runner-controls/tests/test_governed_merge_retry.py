#!/usr/bin/env python3
# ci-pattern asset tests — reintento de las escrituras (DysTelefonica/team-skills#1592)
"""Las escrituras idempotentes de los gates reintentan un 5xx; las no idempotentes, no.

El merge gobernado reintenta el ``PUT .../merge`` con ``sha`` y el ``PATCH`` que
cierra una issue con la regla de ``assets/gh-api-retry/casos.tsv``. El ``POST``
de un comentario o de una issue de incidente hace un solo intento: repetido
tras un 5xx que sí se aplicó, duplicaría el efecto. El stub pone
``GH_API_RETRY_ESPERA_S=0`` (``stub_env``): ningún test duerme de verdad."""
import json
import tempfile
import unittest
from pathlib import Path

from test_governed_merge import base_state, contract
from test_governed_merge import run_gate as run_merge
from test_governed_merge_close import ref, tip_state
from test_push_compliance import REPO, SHA, push_event
from test_push_compliance import base_state as push_state
from test_push_compliance import contract as push_contract
from test_push_compliance import run_gate as run_push


class WriteRetryTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def stored(self) -> dict:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))

    def calls(self, method: str, suffix: str) -> int:
        return sum(1 for call in self.stored()["calls"]
                   if call.startswith(f"{method} ") and call.endswith(suffix))

    def test_a_transient_503_on_the_merge_put_is_retried(self) -> None:
        state = base_state()
        state["merge_transient"] = 1
        proc = run_merge(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.calls("PUT", f"repos/{REPO}/pulls/7/merge"), 2)
        self.assertEqual(len(self.stored()["merges"]), 1)
        self.assertIn(f"AVISO: PUT repos/{REPO}/pulls/7/merge fallo (intento 1 de 4)", proc.stderr)

    def test_a_merge_put_that_keeps_failing_fails_closed_after_four_attempts(self) -> None:
        state = base_state()
        state["merge_transient"] = 99
        proc = run_merge(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.calls("PUT", f"repos/{REPO}/pulls/7/merge"), 4)
        self.assertEqual(self.stored()["merges"], [])
        self.assertIn("the merge call failed (fail closed): HTTP 503", proc.stdout + proc.stderr)
        self.assertIn("(falló tras 4 intentos)", proc.stdout + proc.stderr)

    def test_a_rejected_merge_put_is_not_retried(self) -> None:
        state = base_state()
        state["merge_rejected"] = True
        proc = run_merge(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.calls("PUT", f"repos/{REPO}/pulls/7/merge"), 1)
        self.assertNotIn("AVISO", proc.stderr)

    def test_a_transient_503_on_the_closing_patch_is_retried(self) -> None:
        state = tip_state()
        state["closing_refs"] = [ref(305)]
        state["close_transient"] = 1
        proc = run_merge(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.calls("PATCH", f"repos/{REPO}/issues/305"), 2)
        self.assertEqual(self.stored()["closed_issues"],
                         [{"number": 305, "state": "closed", "state_reason": "completed"}])
        self.assertIn("CLOSED: #305", proc.stdout)

    def test_a_503_on_the_closing_comment_post_is_not_retried(self) -> None:
        state = tip_state()
        state["closing_refs"] = [ref(305)]
        state["comment_error"] = True
        proc = run_merge(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.calls("POST", f"repos/{REPO}/issues/305/comments"), 1)
        self.assertIn("WARNING: closed #305 but could not comment: HTTP 503", proc.stderr)

    def test_a_503_on_the_incident_issue_post_is_not_retried(self) -> None:
        state = push_state()
        state["pulls"][SHA] = []
        state["issue_post_error"] = True
        proc = run_push(self.root, push_event(), push_contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.calls("POST", f"repos/{REPO}/issues"), 1)
        self.assertEqual(self.stored()["created_issues"], [])
        self.assertIn("the incident issue could not be opened (fail closed): HTTP 503",
                      proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main()
