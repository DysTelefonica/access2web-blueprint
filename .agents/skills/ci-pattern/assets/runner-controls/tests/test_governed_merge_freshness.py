#!/usr/bin/env python3
# ci-pattern asset tests — governed merge freshness as a contract parameter (Refs #1397)
"""#1397: el CI corre una vez, en la PR; la frescura respecto de la base es un
parametro del consumer (``merge_freshness`` en ``host-contract.json``).

- Ausente u ``"off"`` (por defecto): un PR atrasado con los checks requeridos en
  verde se fusiona; la red de seguridad es el CI post-merge despachado
  (``post_merge_dispatch``) y el nightly. El arbol integrado distinto del head
  no pone el job en rojo: lo verifica ese CI post-merge.
- ``"required"``: el comportamiento de #787 (atrasado -> 1; drift -> 5).
- El solape de workflows (#783) se rechaza siempre.
- Sin ``post_merge_dispatch`` no hay red post-merge: el atraso sigue bloqueando.
- Un valor desconocido falla cerrado (2).
"""
import json
import tempfile
import unittest
from pathlib import Path

from _github_stub import MERGED_SHA
from test_governed_merge import SHA, TREE_MATCH, commit_obj, run_gate
from test_governed_merge_post_dispatch import contract_with, with_workflow

TREE_BASE_MOVED = "9" * 40


def behind_state(base_changes: list[str], pr_files: list[str]) -> dict:
    """PR atrasado `len(base_changes)` commits; el merge integra los cambios de la base."""
    state = with_workflow("ci", declares=True)
    state["compares"] = {
        f"main...{SHA}": {"behind_by": len(base_changes),
                          "files": [{"filename": name} for name in pr_files]},
        f"{SHA}...main": {"files": [{"filename": name} for name in base_changes]}}
    state["commits"][MERGED_SHA] = commit_obj(MERGED_SHA, TREE_BASE_MOVED)
    return state


def freshness_contract(value: object = None) -> dict:
    data = contract_with(["ci"])
    if value is not None:
        data["merge_freshness"] = value
    return data


class MergeFreshnessTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def stored(self) -> dict:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))

    def test_behind_pr_with_green_checks_merges_by_default(self) -> None:
        # S1: atrasado, checks en verde, sin clave -> se fusiona y se despacha el CI post-merge.
        summary = self.root / "summary.md"
        proc = run_gate(self.root, freshness_contract(),
                        behind_state(["docs/guide.md"], ["src/app.php"]), summary=str(summary))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("BLOCKED", proc.stdout)
        self.assertIn(f"MERGED: PR #7 merged by the governed path as {MERGED_SHA}", proc.stdout)
        self.assertTrue(self.stored()["merges"])
        self.assertTrue(any("/dispatches" in call for call in self.stored()["calls"]))
        # S3: el arbol integrado no es el del head, y el veredicto lo atribuye al CI post-merge.
        text = summary.read_text(encoding="utf-8")
        for expected in (TREE_BASE_MOVED, TREE_MATCH,
                         "the integrated tree includes base commits the PR did not carry; "
                         "the post-merge CI on 'main' verifies it"):
            self.assertIn(expected, text)
        self.assertNotIn("is NOT the tree CI reviewed", text + proc.stdout + proc.stderr)

    def test_explicit_off_is_the_default(self) -> None:
        proc = run_gate(self.root, freshness_contract("off"),
                        behind_state(["docs/guide.md"], ["src/app.php"]))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertTrue(self.stored()["merges"])

    def test_workflow_overlap_is_still_refused_with_its_message(self) -> None:
        # S2: el unico rechazo por frescura que queda, con su mensaje actual.
        proc = run_gate(self.root, freshness_contract(),
                        behind_state([".github/workflows/ci.yml"], [".github/workflows/ci.yml"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("PR #7 and 'main' both changed .github/workflows/ci.yml: merging would create "
                      "new workflow content, which the workflow token cannot do (no 'workflows' "
                      "permission); update the PR branch onto 'main', let CI go green and run the "
                      "governed merge again", proc.stdout)
        self.assertIn("VERDICT: PR #7 not merged: contract not satisfied", proc.stdout)
        self.assertFalse(self.stored()["merges"])

    def test_required_keeps_the_787_refusal_and_drift_exit(self) -> None:
        # S6: consumers whose deploy gate compares trees (APAP_WEB) opt in.
        proc = run_gate(self.root, freshness_contract("required"),
                        behind_state(["docs/guide.md"], ["src/app.php"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("PR #7 is behind 'main' by 1 commit(s): the merged tree would differ from the "
                      "tree CI reviewed and the contract sets merge_freshness to 'required'; update "
                      "the PR branch onto 'main', let CI go green and run the governed merge again",
                      proc.stdout)
        self.assertFalse(self.stored()["merges"])

    def test_required_still_reports_drift_as_5(self) -> None:
        state = behind_state([], ["src/app.php"])  # al dia en el check, la base se mueve antes del PUT
        proc = run_gate(self.root, freshness_contract("required"), state)
        self.assertEqual(proc.returncode, 5, proc.stdout + proc.stderr)
        self.assertIn("the integrated tree is NOT the tree CI reviewed", proc.stdout)

    def test_without_post_merge_dispatch_a_behind_pr_is_still_blocked(self) -> None:
        # S3: sin CI post-merge declarado no hay red de seguridad que verifique el arbol fusionado.
        data = freshness_contract()
        del data["post_merge_dispatch"]
        proc = run_gate(self.root, data, behind_state(["docs/guide.md"], ["src/app.php"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("PR #7 is behind 'main' by 1 commit(s): the merged tree would differ from the "
                      "tree CI reviewed and the contract declares no post_merge_dispatch to verify "
                      "it after the merge; update the PR branch onto 'main', let CI go green and run "
                      "the governed merge again", proc.stdout)
        self.assertFalse(self.stored()["merges"])

    def test_up_to_date_pr_merges_with_the_matching_verdict(self) -> None:
        state = behind_state([], ["src/app.php"])
        state["commits"][MERGED_SHA] = commit_obj(MERGED_SHA, TREE_MATCH)
        proc = run_gate(self.root, freshness_contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("verdict: the integrated tree is the tree CI reviewed", proc.stdout)

    def test_unknown_value_fails_closed_before_any_api_call(self) -> None:
        for value in ("strict", True, 1):
            with self.subTest(value=value):
                proc = run_gate(self.root, freshness_contract(value),
                                behind_state(["docs/guide.md"], ["src/app.php"]))
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertIn("merge_freshness must be 'off' or 'required'", proc.stderr)
                self.assertEqual(self.stored()["calls"], [])


if __name__ == "__main__":
    unittest.main()
