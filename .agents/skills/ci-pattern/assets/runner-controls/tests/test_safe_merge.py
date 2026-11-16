#!/usr/bin/env python3
# ci-pattern asset tests -- safe merge, parte 1: validaciones (#426)
"""Validaciones de estado y CI del gate safe-merge.

Cubre los criterios del issue #426 sobre el estado del PR antes del
merge: PR abierto, base == default_branch, checks verdes y
required-checks presentes. El atraso respecto de la base solo se rechaza
con ``--require-fresh``: el CI corre una vez, en la PR (#1397). Los tests de chain/closing y
API errors viven en test_safe_merge_chain.py para mantener este
archivo bajo el presupuesto de revision.

Shim y fixtures compartidos en _shim.py.
"""
import json
import os
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _shim import (
    PATH_TESTS_SKIP,
    SHA,
    BaseSafeMergeTest,
    base_state,
    run_gate,
)


class ValidationsTest(BaseSafeMergeTest):
    """#426 parte 1: el gate verifica el estado del PR antes del merge."""

    def test_pr_behind_is_denied_with_require_fresh(self) -> None:
        state = base_state()
        state["compare"]["behind_by"] = 3
        proc = run_gate(self.root, state, extra=("--require-fresh",))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("DENIED: PR #7 is behind by 3 commit(s) against 'main'; run `git fetch "
                      "origin && git merge origin/main`, wait for CI green, then retry.", proc.stdout)

    def test_pr_behind_with_green_checks_merges_by_default(self) -> None:
        # #1397: sin --require-fresh, un PR atrasado y en verde se fusiona.
        state = base_state()
        state["compare"]["behind_by"] = 3
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)
        self.assertNotIn("DENIED", proc.stdout)
        self.assertIn("MERGED: PR #7 merged by safe-merge as", proc.stdout)

    def test_red_check_is_denied(self) -> None:
        state = base_state()
        state["check_runs"][SHA][0]["conclusion"] = "failure"
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("check", proc.stdout.lower())

    def test_non_default_base_is_denied(self) -> None:
        state = base_state()
        state["default_branch"] = "main"
        state["pr_details"]["7"]["base"]["ref"] = "develop"
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("base", proc.stdout.lower())

    def test_happy_path_merges_and_cleans(self) -> None:
        proc = run_gate(self.root, base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)
        # #1491: la rama remota se conserva siempre (invariante de la flota).
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["deleted_branches"], [])
        self.assertIn("KEEP: remote branch 'chore/test' is kept: the fleet never deletes "
                      "remote branches", proc.stdout)

    @unittest.skipIf(os.name == "nt", PATH_TESTS_SKIP)
    def test_without_explicit_gh_the_gh_on_path_still_merges(self) -> None:
        # #1609: sin CI_PATTERN_GH el comportamiento de produccion no cambia.
        proc = run_gate(self.root, base_state(), via_path=True)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)

    def test_check_only_exits_zero(self) -> None:
        proc = run_gate(self.root, base_state(), check_only=True)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("check-only", proc.stdout)

    def test_missing_required_check_denied(self) -> None:
        # El check existe pero no es el requerido por el gate.
        state = base_state()
        proc = run_gate(self.root, state, required_check=["other-check"])
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)

    def test_open_state_required(self) -> None:
        state = base_state()
        state["pr_details"]["7"]["state"] = "closed"
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("open", proc.stdout.lower())



def _run(run_id: int, conclusion: str) -> dict:
    return {"id": run_id, "name": "unit", "conclusion": conclusion, "status": "completed"}


class CancelledTwinTest(BaseSafeMergeTest):
    """#1598: un check-run ``cancelled`` con gemelo del mismo nombre no es veredicto del SHA.

    Por cada nombre decide el check-run mas reciente no cancelado; si todos estan
    cancelados, decide el cancelado mas reciente y el merge se bloquea (como #1591)."""

    def gate(self, *runs: dict):
        state = base_state()
        state["check_runs"][SHA] = list(runs)
        return run_gate(self.root, state, required_check=["unit"])

    def test_old_run_cancelled_by_the_new_one_merges_when_the_new_one_is_green(self) -> None:
        # Linea temporal: el run viejo (id menor) se cancela al llegar el nuevo (id mayor).
        proc = self.gate(_run(1, "cancelled"), _run(2, "success"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)
        self.assertNotIn("DENIED", proc.stdout)

    def test_newer_cancelled_twin_of_a_green_run_merges(self) -> None:
        proc = self.gate(_run(1, "success"), _run(2, "cancelled"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)
        self.assertNotIn("DENIED", proc.stdout)

    def test_cancelled_alone_is_denied(self) -> None:
        proc = self.gate(_run(1, "cancelled"))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("check 'unit': status=completed, conclusion=cancelled", proc.stdout)

    def test_all_twins_cancelled_is_denied(self) -> None:
        proc = self.gate(_run(1, "cancelled"), _run(2, "cancelled"))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("check 'unit': status=completed, conclusion=cancelled", proc.stdout)

    def test_newer_cancelled_twin_does_not_hide_an_older_failure(self) -> None:
        proc = self.gate(_run(1, "failure"), _run(2, "cancelled"))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("check 'unit': status=completed, conclusion=failure", proc.stdout)


if __name__ == "__main__":
    unittest.main()
