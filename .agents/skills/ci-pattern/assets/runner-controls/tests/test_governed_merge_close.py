#!/usr/bin/env python3
# ci-pattern asset tests — governed merge closes linked issues (DysTelefonica/team-skills#800)
"""After a governed merge the linked issues are closed explicitly.

Merges made with the workflow `GITHUB_TOKEN` do not fire GitHub's implicit
closing of `Closes #N` issues (0 of 13 in APAP and CADETE, #800), so the
governed path closes what `closingIssuesReferences` reports. The merge already
happened by then: a failure to close is exit 3, never the fail-closed 2."""
import json
import tempfile
import unittest
from pathlib import Path

from test_governed_merge import base_state, contract, run_gate
from test_push_compliance import REPO


def ref(number: int, state: str = "OPEN", repo: str = REPO) -> dict:
    return {"number": number, "state": state, "repository": {"nameWithOwner": repo}}


class GovernedMergeCloseTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def stored(self) -> dict:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))

    def test_merge_closes_the_open_linked_issue_as_completed(self) -> None:
        state = base_state()
        state["closing_refs"] = [ref(305)]
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        stored = self.stored()
        self.assertEqual(stored["closed_issues"],
                         [{"number": 305, "state": "closed", "state_reason": "completed"}])
        self.assertIn("PR #7", stored["comments"][0]["body"])
        self.assertIn("CLOSED: #305", proc.stdout)

    def test_closed_foreign_and_unlinked_issues_are_left_alone(self) -> None:
        state = base_state()
        state["closing_refs"] = [ref(10, "CLOSED"), ref(11, repo="DysTelefonica/otro")]
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("closed_issues", self.stored())
        self.assertNotIn("comments", self.stored())

    def test_pr_without_closing_references_closes_nothing(self) -> None:
        proc = run_gate(self.root, contract(), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("closed_issues", self.stored())

    def test_blocked_merge_does_not_look_up_or_close_anything(self) -> None:
        state = base_state()
        state["closing_refs"] = [ref(305)]
        state["pr_details"]["7"]["draft"] = True
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        stored = self.stored()
        self.assertNotIn("closed_issues", stored)
        self.assertFalse([call for call in stored["calls"] if "graphql" in call])

    def test_base_other_than_the_default_branch_closes_nothing(self) -> None:
        # GitHub only closes on the default branch; a chain intermediate targets another base.
        state = base_state()
        state["closing_refs"] = [ref(305)]
        state["default_branch"] = "trunk"
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("closed_issues", self.stored())

    def test_failure_to_close_after_the_merge_is_exit_3_naming_the_issue(self) -> None:
        state = base_state()
        state["closing_refs"] = [ref(305)]
        state["close_error"] = True
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 3, proc.stdout + proc.stderr)
        self.assertTrue(self.stored()["merges"])
        output = proc.stdout + proc.stderr
        for expected in ("MERGED", "#305", "issues: write"):
            self.assertIn(expected, output)

    def test_failed_lookup_after_the_merge_is_exit_3(self) -> None:
        for flag in ("graphql_error", "graphql_errors"):
            with self.subTest(flag=flag):
                state = base_state()
                state[flag] = True
                proc = run_gate(self.root, contract(), state)
                self.assertEqual(proc.returncode, 3, proc.stdout + proc.stderr)
                self.assertTrue(self.stored()["merges"])
                self.assertIn("could not list", proc.stdout + proc.stderr)

    def test_patch_422_when_issue_already_closed_is_treated_as_success(self) -> None:
        # #835: PATCH /issues/N puede devolver 422 si el cierre automático de
        # GitHub por la palabra Closes ya cerró la issue (carrera con el PATCH
        # explícito del gate). Re-read y aceptar: la pista sobre `issues: write`
        # era engañosa. Salida: exit 0, mensaje CLOSED visible, sin exit 3.
        state = base_state()
        state["closing_refs"] = [ref(305)]
        state["race_after_close"] = True
        # El re-read por defecto del stub devuelve `closed` (la carrera ya
        # cerró la issue). Se valida con `state["issue_states"][305]` el caso
        # «todavía abierta» en el test siguiente.
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        output = proc.stdout + proc.stderr
        self.assertIn("CLOSED", output)
        self.assertIn("#305", output)
        # El 422 ya no se reporta como fallo: el re-read confirmó el cierre.
        self.assertNotIn("HTTP 422", output)
        self.assertNotIn("issues: write", output)

    def test_patch_422_when_issue_still_open_still_fails_closed(self) -> None:
        # #835: si tras el PATCH 422 el re-read dice `state == "open"`, no
        # es una carrera: el merge sale con exit 3 como antes y conserva la
        # pista accionable cuando aplique (no es este caso: el problema es
        # otro, no los permisos).
        state = base_state()
        state["closing_refs"] = [ref(305)]
        state["race_after_close"] = True
        state["issue_states"] = {305: {"state": "open", "state_reason": None}}
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 3, proc.stdout + proc.stderr)
        self.assertTrue(self.stored()["merges"])
        output = proc.stdout + proc.stderr
        self.assertIn("#305", output)
        self.assertIn("HTTP 422", output)


if __name__ == "__main__":
    unittest.main()
