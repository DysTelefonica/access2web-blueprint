#!/usr/bin/env python3
# ci-pattern asset tests -- safe merge, parte 1: validaciones (#426)
"""Validaciones de estado y CI del gate safe-merge.

Cubre los criterios del issue #426 sobre el estado del PR antes del
merge: PR abierto, base == default_branch, behind_by == 0, checks
verdes y required-checks presentes. Los tests de chain/closing y
API errors viven en test_safe_merge_chain.py para mantener este
archivo bajo el presupuesto de revision.

Shim y fixtures compartidos en _shim.py.
"""
import json
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _shim import (
    SHA,
    BaseSafeMergeTest,
    base_state,
    run_gate,
)


class ValidationsTest(BaseSafeMergeTest):
    """#426 parte 1: el gate verifica el estado del PR antes del merge."""

    def test_pr_behind_is_denied(self) -> None:
        state = base_state()
        state["compare"]["behind_by"] = 3
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("behind", proc.stdout)

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
        # Cleanup: la rama remota se borra.
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertTrue(any("git/refs/heads/chore/test" in c
                            for c in state["deleted_branches"]),
                        state["deleted_branches"])

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


if __name__ == "__main__":
    unittest.main()
