#!/usr/bin/env python3
# ci-pattern asset tests — governed merge (DysTelefonica/team-skills#303)
"""Governed merge suite (black-box): only a PR green on every contract
check, with exactly one type label, within the review budget and without
Closes in an intermediate, is merged through the API (exit 0); any
contract violation blocks the merge (exit 1) and an unscannable
environment fails closed (exit 2). The API is a local HTTP stub with real shapes."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from _github_stub import github_stub, stub_env
from test_push_compliance import CHECK, REPO, SHA

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


def run_gate(root: Path, contract_data: dict, state: dict, *extra: str) -> subprocess.CompletedProcess:
    (root / "contract.json").write_text(json.dumps(contract_data), encoding="utf-8")
    with github_stub(root / "state.json", state) as base:
        return subprocess.run([sys.executable, str(ASSET), "--repo", REPO, "--pr", "7",
                               "--contract", str(root / "contract.json"), *extra],
                              capture_output=True, text=True, env=stub_env(base), timeout=60)


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

    def test_oversized_pr_is_not_re_evaluated(self) -> None:
        # #787: el tamano lo mide pr-size (que excluye lo vendorizado y exige la etiqueta);
        # el merge no repite la cuenta. APAP #2122: 19.786 lineas vendorizadas, bloqueado.
        state = base_state()
        state["pr_details"]["7"].update(additions=19000, deletions=786)
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)

    def test_legacy_max_lines_flag_is_accepted_and_ignored(self) -> None:
        # Los workflows de los consumidores aun pasan --max-lines; quitarlo romperia su merge.
        state = base_state()
        state["pr_details"]["7"].update(additions=380, deletions=60)
        proc = run_gate(self.root, contract(), state, "--max-lines", "400")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)

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

    def workflow_state(self, base_changes: list[str]) -> dict:
        """PR que cambia ci.yml y una base que cambió `base_changes` desde la bifurcacion."""
        state = base_state()
        state["compares"] = {
            f"main...{SHA}": {"behind_by": len(base_changes),
                              "files": [{"filename": ".github/workflows/ci.yml"},
                                        {"filename": "README.md"}]},
            f"{SHA}...main": {"files": [{"filename": name} for name in base_changes]}}
        return state

    def test_workflow_overlap_with_the_base_is_blocked_before_the_merge(self) -> None:
        # #2114: la base tambien cambio ci.yml; el merge crearia un blob de workflow
        # nuevo y GITHUB_TOKEN no puede -> veredicto accionable (1), no 403 opaco (2).
        proc = run_gate(self.root, contract(), self.workflow_state([".github/workflows/ci.yml"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn(".github/workflows/ci.yml", proc.stdout)
        self.assertIn("update the PR branch", proc.stdout)

    def test_workflow_change_on_an_up_to_date_branch_still_merges(self) -> None:
        # #2117 y #2110: tocan ci.yml pero la base no se movio -> la cabeza ya trae el fichero final.
        proc = run_gate(self.root, contract(), self.workflow_state([]))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)

    def test_branch_behind_the_base_is_blocked_before_the_merge(self) -> None:
        # #787 S4: con la base movida el arbol fusionado deja de ser el que reviso el CI y el
        # deploy lo rechaza ("main moved"); se bloquea antes, con la accion a tomar.
        proc = run_gate(self.root, contract(), self.workflow_state(["docs/guide.md"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn("behind 'main' by 1 commit(s)", proc.stdout)
        self.assertIn("update the PR branch", proc.stdout)

    def test_compare_without_behind_by_fails_closed(self) -> None:
        state = base_state()
        state["compare"] = {"files": []}
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)

    def test_pr_without_workflow_changes_skips_the_overlap_lookup(self) -> None:
        proc = run_gate(self.root, contract(), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(sum("/compare/" in call for call in state["calls"]), 1)

    def test_malformed_compare_fails_closed(self) -> None:
        state = self.workflow_state([])
        state["compares"][f"main...{SHA}"] = {"files": "nope"}
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)

    def test_forbidden_merge_names_github_reason_and_the_workflows_hint(self) -> None:
        state = base_state()
        state["merge_forbidden"] = "refusing to allow a GitHub App to update workflow without `workflows` permission"
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        for expected in ("HTTP 403", "refusing to allow", "REQ-1234", "'workflows' permission"):
            self.assertIn(expected, proc.stderr)

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

    def test_required_check_that_appears_later_is_waited_for(self) -> None:
        # #2107: el merge se lanzo antes de que el CI publicara `required`; esperar, no fallar.
        state = base_state()
        green = [{"id": 1, "name": CHECK, "conclusion": "success", "status": "completed"}]
        pending = [{"id": 1, "name": CHECK, "conclusion": None, "status": "in_progress"}]
        state["check_runs_script"] = {SHA: [[], pending, green]}
        proc = run_gate(self.root, contract(), state, "--wait-seconds", "30", "--poll-seconds", "0.05")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)

    def test_required_check_still_missing_after_the_wait_blocks(self) -> None:
        state = base_state()
        state["check_runs"][SHA] = []
        proc = run_gate(self.root, contract(), state, "--wait-seconds", "0.3", "--poll-seconds", "0.05")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)
        self.assertIn(f'required check "{CHECK}": missing', proc.stdout)

    def test_red_check_does_not_wait(self) -> None:
        state = base_state()
        state["check_runs"][SHA] = [{"id": 1, "name": CHECK, "conclusion": "failure", "status": "completed"}]
        proc = run_gate(self.root, contract(), state, "--wait-seconds", "30", "--poll-seconds", "0.05")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        stored = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(sum("/check-runs" in call for call in stored["calls"]), 1)

    def test_negative_wait_is_rejected(self) -> None:
        proc = run_gate(self.root, contract(), base_state(), "--wait-seconds", "-1")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("--wait-seconds must be", proc.stderr)


if __name__ == "__main__":
    unittest.main()
