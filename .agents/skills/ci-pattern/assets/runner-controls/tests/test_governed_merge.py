#!/usr/bin/env python3
# ci-pattern asset tests — governed merge (DysTelefonica/team-skills#303)
"""Governed merge suite (black-box): only a PR green on every contract
check, with exactly one type label, within the review budget and without
Closes in an intermediate, is merged through the API (exit 0); any
contract violation blocks the merge (exit 1) and an unscannable
environment fails closed (exit 2). ``gh`` simulated with real API shapes."""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from test_push_compliance import CHECK, GH_SHIM, REPO, SHA

ASSET = Path(__file__).resolve().parent.parent / "check_governed_merge.py"


def pull(**over: object) -> dict:
    base: dict = {"number": 7, "state": "open", "mergeable": True, "draft": False,
                  "user": {"login": "ardelperal"}, "head": {"sha": SHA}, "base": {"ref": "main"},
                  "title": "feat: change", "body": "Refs #305",
                  "labels": [{"name": "type:bug"}, {"name": "chain:partial"}],
                  "additions": 10, "deletions": 5, "changed_files": 2}
    base.update(over)
    return base


def contract() -> dict:
    return {"contract_version": 1, "protected_branches": ["main"],
            "required_checks": [{"name": CHECK, "class": "documented-only"}]}


def base_state() -> dict:
    return {"pr_details": {"7": pull()},
            "check_runs": {SHA: [{"id": 1, "name": CHECK, "conclusion": "success", "status": "completed"}]},
            "created_issues": [], "calls": [], "merges": []}


def run_gate(root: Path, contract_data: dict, state: dict) -> subprocess.CompletedProcess:
    (root / "contract.json").write_text(json.dumps(contract_data), encoding="utf-8")
    (root / "state.json").write_text(json.dumps(state), encoding="utf-8")
    shim = root / "bin" / "gh"
    shim.parent.mkdir(exist_ok=True)
    shim.write_text(GH_SHIM.format(state_path=str(root / "state.json")), encoding="utf-8")
    shim.chmod(0o755)
    env = {**os.environ, "PATH": f"{root / 'bin'}{os.pathsep}{os.environ.get('PATH', '')}"}
    return subprocess.run([sys.executable, str(ASSET), "--repo", REPO, "--pr", "7",
                           "--contract", str(root / "contract.json")],
                          capture_output=True, text=True, env=env, timeout=60)


class GovernedMergeTests(unittest.TestCase):
    """#303 control 2: el merge sale por la API solo con el contrato verde."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def assert_merged(self, proc: subprocess.CompletedProcess, expected: bool) -> None:
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(bool(state["merges"]), expected, proc.stdout + proc.stderr)

    def test_green_pr_is_merged_via_api(self) -> None:
        proc = run_gate(self.root, contract(), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)
        self.assertIn("MERGED", proc.stdout)

    def test_red_check_blocks_the_merge(self) -> None:
        state = base_state()
        state["check_runs"][SHA] = [{"id": 1, "name": CHECK, "conclusion": "failure", "status": "completed"}]
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn(CHECK, proc.stdout + proc.stderr)

    def test_missing_type_label_blocks_the_merge(self) -> None:
        state = base_state()
        state["pr_details"]["7"]["labels"] = [{"name": "chain:partial"}]
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("type", proc.stdout + proc.stderr)

    def test_oversized_pr_blocks_the_merge(self) -> None:
        state = base_state()
        state["pr_details"]["7"].update(additions=380, deletions=60)
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("400", proc.stdout + proc.stderr)

    def test_closes_in_intermediate_pr_blocks_the_merge(self) -> None:
        state = base_state()
        state["pr_details"]["7"]["body"] = "Refs #303\nCloses #303"
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("Closes", proc.stdout + proc.stderr)

    def test_concurrent_push_between_check_and_merge_fails_closed(self) -> None:
        # TOCTOU: el head verificado cambia antes del PUT; el sha empujado
        # hace que la API responda 409 -> exit 2, nunca un merge sin verificar.
        state = base_state()
        state["head_changed"] = True
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)

    def test_size_exception_allows_oversized_pr(self) -> None:
        # HR-8: el campo size-exception-reason ES la excepcion del presupuesto.
        state = base_state()
        state["pr_details"]["7"].update(
            additions=380, deletions=60,
            body="Refs #303\nsize-exception-reason: borrado completo de un fichero\n")
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)

    def test_invalid_size_exception_blocks_the_merge(self) -> None:
        # Vacio o duplicado: la excepcion no es dato -> el tamano bloquea.
        for body in ("Refs #303\nsize-exception-reason:\n",
                     "Refs #303\nsize-exception-reason: a\nsize-exception-reason: b\n"):
            state = base_state()
            state["pr_details"]["7"].update(additions=380, deletions=60, body=body)
            proc = run_gate(self.root, contract(), state)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assert_merged(proc, False)

    def test_draft_pr_blocks_the_merge(self) -> None:
        state = base_state()
        state["pr_details"]["7"]["draft"] = True
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("draft", proc.stdout + proc.stderr)

    def test_non_protected_base_is_blocked_naming_it(self) -> None:
        # Coherente con el detective: merge solo a rama protegida por el
        # contrato o a eslabon de cadena (chain:partial en el PR).
        state = base_state()
        state["pr_details"]["7"]["base"]["ref"] = "features/suelta"
        state["pr_details"]["7"]["labels"] = [{"name": "type:bug"}]  # punta, sin cadena
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("features/suelta", proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        state = base_state()  # el mismo PR como intermedio de cadena: si cabe
        state["pr_details"]["7"]["base"]["ref"] = "features/suelta"
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_fail_closed_environment(self) -> None:
        state = base_state()
        state["api_error"] = True
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        state = base_state()
        state["merge_rejected"] = True
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail", (proc.stdout + proc.stderr).lower())


if __name__ == "__main__":
    unittest.main()
