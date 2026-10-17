#!/usr/bin/env python3
# ci-pattern asset tests -- safe merge, parte 2: chain y API (#426)
"""Validaciones de chain/closing y errores de API del gate safe-merge.

Cubre los criterios del issue #426 sobre la coherencia de la cadena
(``chain:partial`` exige ``Refs #N`` sin ``Closes``; la punta exige
palabra de cierre) y el fail-closed ante errores de la API de GitHub
(5xx nunca se toma como verde; 409 TOCTOU aborta).

Los tests de validaciones de estado/CI viven en
test_safe_merge.py. Shim y fixtures compartidos en _shim.py.
"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _shim import BaseSafeMergeTest, base_state, run_gate


class ChainAndApiTest(BaseSafeMergeTest):
    """#426 parte 2: coherencia de cadena y fail-closed ante errores de API."""

    def test_chain_partial_with_closing_is_denied(self) -> None:
        # chain:partial con Closes -> el gate lo niega (mezcla prohibida).
        state = base_state()
        state["pr_details"]["7"]["body"] = (
            "<!-- type:feature chain:partial Closes #426 -->\n\nbody"
        )
        state["pr_details"]["7"]["labels"] = [{"name": "chain:partial"}]
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("chain", proc.stdout.lower())

    def test_chain_partial_without_refs_is_denied(self) -> None:
        # chain:partial sin Refs #N -> falta el enlace a la issue.
        state = base_state()
        state["pr_details"]["7"]["body"] = (
            "<!-- type:feature chain:partial -->\n\nbody sin refs"
        )
        state["pr_details"]["7"]["labels"] = [{"name": "chain:partial"}]
        state["pr_details"]["7"]["body"] = state["pr_details"]["7"]["body"].replace(
            "body", "Closes #426"
        )
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("refs", proc.stdout.lower())

    def test_chain_tip_without_closing_is_denied(self) -> None:
        # Punta (sin chain:partial) sin palabra de cierre -> el gate lo niega.
        state = base_state()
        state["pr_details"]["7"]["body"] = "feat: sin palabra de cierre"
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("closing", proc.stdout.lower())

    def test_chain_partial_with_refs_ok(self) -> None:
        # chain:partial con Refs #N (sin Closes) -> OK.
        state = base_state()
        state["pr_details"]["7"]["body"] = (
            "<!-- type:feature chain:partial Refs #426 -->\n\nbody"
        )
        state["pr_details"]["7"]["labels"] = [{"name": "chain:partial"}]
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)

    def test_api_error_fails_closed(self) -> None:
        # 5xx de GitHub -> exit 2, no se toma como verde.
        state = base_state()
        state["api_error"] = True
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("api", proc.stdout.lower() + proc.stderr.lower())

    def test_head_changed_409_fails_closed(self) -> None:
        # TOCTOU: la cabeza cambio entre verificacion y merge -> 409.
        state = base_state()
        state["head_changed"] = True
        proc = run_gate(self.root, state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("409", proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main()
