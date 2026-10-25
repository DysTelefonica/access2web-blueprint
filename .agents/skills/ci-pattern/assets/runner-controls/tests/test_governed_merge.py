#!/usr/bin/env python3
# ci-pattern asset tests — governed merge (DysTelefonica/team-skills#303)
"""Governed merge suite (black-box): only a PR green on every contract
check, with exactly one type label, within the review budget and without
Closes in an intermediate, is merged through the API (exit 0); any
contract violation blocks the merge (exit 1) and an unscannable
environment fails closed (exit 2). The API is a local HTTP stub with real shapes.

#1265: tras el merge, el job mide el arbol integrado (el del commit de merge
frente al del head verificado) y publica el veredicto en stdout y en
``$GITHUB_STEP_SUMMARY``: arboles iguales, 0; arboles distintos, 5; una
medicion que no se puede hacer, 2 con la causa nombrada."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from _github_stub import MERGED_SHA, github_stub, stub_env
from test_push_compliance import CHECK, REPO, SHA

ASSET = Path(__file__).resolve().parent.parent / "check_governed_merge.py"

# #1265: shas de arbol de 40 hex, la forma que acepta el lector compartido.
TREE_MATCH = "1a2b3c4d5e6f708192a3b4c5d6e7f8091a2b3c4d"
TREE_DRIFT = "5e6f708192a3b4c5d6e7f8091a2b3c4d5e6f7081"
# #1265: primer padre que NO es la base (un rebase-merge): el veredicto sale
# igual, el control no se salta.
REBASED_PARENT = "f" * 40


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


def commit_obj(sha: str, tree: str) -> dict:
    """Commit con la forma real de la API: el arbol vive en ``commit.tree.sha`` (#928)."""
    return {"sha": sha, "commit": {"tree": {"sha": tree}}}


def base_state() -> dict:
    return {"pr_details": {"7": pull()},
            "check_runs": {SHA: [{"id": 1, "name": CHECK, "conclusion": "success", "status": "completed"}]},
            # #1265: el veredicto del arbol integrado lee los dos commits; por
            # defecto sus arboles coinciden, que es el caso sano.
            "commits": {MERGED_SHA: commit_obj(MERGED_SHA, TREE_MATCH),
                        SHA: commit_obj(SHA, TREE_MATCH)},
            "created_issues": [], "calls": [], "merges": []}


def run_gate(root: Path, contract_data: dict, state: dict, *extra: str,
             summary: str = "") -> subprocess.CompletedProcess:
    """Corre el asset contra el stub; ``summary`` fija ``GITHUB_STEP_SUMMARY``.

    Sin ``summary`` la variable no se define: una corrida local o un test no
    tienen resumen del job, y el veredicto queda solo en stdout.
    """
    (root / "contract.json").write_text(json.dumps(contract_data), encoding="utf-8")
    with github_stub(root / "state.json", state) as base:
        env = stub_env(base)
        env.pop("GITHUB_STEP_SUMMARY", None)
        if summary:
            env["GITHUB_STEP_SUMMARY"] = summary
        return subprocess.run([sys.executable, str(ASSET), "--repo", REPO, "--pr", "7",
                               "--contract", str(root / "contract.json"), *extra],
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

    def test_missing_type_label_does_not_block_the_merge(self) -> None:
        # R1 (DysTelefonica/team-skills#1093): la etiqueta `type:*` del PR ya no
        # es un gate; la llevan los agentes por costumbre.
        state = base_state()
        state["pr_details"]["7"]["labels"] = [{"name": "chain:partial"}]
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)

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


class IntegratedTreeEvidenceTests(unittest.TestCase):
    """#1265: tras el merge, el job publica la evidencia del arbol integrado.

    El veredicto compara el ARBOL del commit de merge con el ARBOL del head
    verificado (no los commits: el metodo de merge del host no puede cambiar el
    veredicto) y lo escribe en stdout y en ``$GITHUB_STEP_SUMMARY``.
    """

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def assert_merged(self, proc: subprocess.CompletedProcess, expected: bool) -> None:
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(bool(state["merges"]), expected, proc.stdout + proc.stderr)

    def test_arboles_iguales_sale_0_y_el_resumen_nombra_los_cuatro_sha(self) -> None:
        # Criterio 1: el arbol del commit de merge es el del head verificado ->
        # 0, y el resumen nombra los dos commits y los dos arboles.
        summary = self.root / "summary.md"
        proc = run_gate(self.root, contract(), base_state(), summary=str(summary))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        text = summary.read_text(encoding="utf-8")
        for expected in (MERGED_SHA, SHA, TREE_MATCH):
            self.assertIn(expected, text)
        self.assertIn("the integrated tree is the tree CI reviewed", text)

    def test_arboles_distintos_sale_5_y_el_resumen_nombra_el_desajuste(self) -> None:
        # Criterio 2: los arboles difieren -> el job se pone en rojo (5) y el
        # resumen dice que el arbol integrado NO es el que reviso el CI. El
        # merge ya esta hecho y no se deshace.
        state = base_state()
        state["commits"][MERGED_SHA] = commit_obj(MERGED_SHA, TREE_DRIFT)
        summary = self.root / "summary.md"
        proc = run_gate(self.root, contract(), state, summary=str(summary))
        self.assertEqual(proc.returncode, 5, proc.stdout + proc.stderr)
        text = summary.read_text(encoding="utf-8")
        self.assertIn(TREE_DRIFT, text)
        self.assertIn(TREE_MATCH, text)
        self.assertIn("the integrated tree is NOT the tree CI reviewed", text)
        self.assert_merged(proc, True)

    def test_el_veredicto_se_escribe_en_el_resumen_del_job(self) -> None:
        # Criterio 3: el veredicto va a $GITHUB_STEP_SUMMARY, no solo a stdout.
        summary = self.root / "summary.md"
        proc = run_gate(self.root, contract(), base_state(), summary=str(summary))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        text = summary.read_text(encoding="utf-8")
        self.assertIn("## Integrated tree evidence", text)
        self.assertIn("the integrated tree is the tree CI reviewed", text)
        self.assertIn("the integrated tree is the tree CI reviewed", proc.stdout)
        # Con la variable definida pero la escritura imposible, el job falla
        # cerrado (2) nombrando la causa: no se puede afirmar que se publico.
        proc = run_gate(self.root, contract(), base_state(),
                        summary=str(self.root / "no" / "summary.md"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("GITHUB_STEP_SUMMARY", proc.stderr)

    def test_el_asset_no_crea_ni_muta_check_runs(self) -> None:
        # Criterio 4: ninguna llamada crea o muta un check-run; la cabeza del PR
        # no gana ningun contexto de check.
        proc = run_gate(self.root, contract(), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        stored = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        check_calls = [call for call in stored["calls"] if "check-runs" in call]
        self.assertTrue(check_calls, stored["calls"])
        for call in check_calls:
            self.assertTrue(call.startswith("GET "), call)

    def test_el_arbol_se_lee_de_commit_tree_sha_y_no_del_tree_de_nivel_raiz(self) -> None:
        # Criterio 5: el `tree.sha` de nivel raiz no es el arbol del commit
        # (regresion de #928): se rechaza, no se acepta en silencio.
        state = base_state()
        state["commits"][MERGED_SHA] = {"sha": MERGED_SHA, "tree": {"sha": TREE_MATCH}, "commit": {}}
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        output = proc.stdout + proc.stderr
        self.assertIn("commit.tree.sha", output)
        self.assertNotIn("the integrated tree is the tree CI reviewed", output)

    def test_el_rebase_merge_se_compara_y_se_reporta_en_vez_de_saltarse(self) -> None:
        # Criterio 6: un commit de merge cuyo primer padre no es la base
        # (rebase-merge) se compara igual: el veredicto sale, no se salta.
        state = base_state()
        state["commits"][MERGED_SHA] = {"sha": MERGED_SHA, "parents": [{"sha": REBASED_PARENT}],
                                        "commit": {"tree": {"sha": TREE_MATCH}}}
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("the integrated tree is the tree CI reviewed", proc.stdout)

    def test_una_lectura_ilegible_no_se_reporta_como_match(self) -> None:
        # Criterio 7: si el commit no se puede leer o no trae la forma esperada,
        # el job falla cerrado (2) nombrando la causa; nunca se informa de una
        # coincidencia que no se ha medido.
        state = base_state()
        del state["commits"][MERGED_SHA]  # la API responde 404
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("HTTP 404", proc.stderr)
        self.assertNotIn("the integrated tree is the tree CI reviewed", proc.stdout + proc.stderr)
        state = base_state()
        state["commits"][MERGED_SHA] = {"sha": MERGED_SHA, "commit": {"tree": {"sha": "no-es-un-sha"}}}
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        output = proc.stdout + proc.stderr
        self.assertIn("commit.tree.sha", output)
        self.assertNotIn("the integrated tree is the tree CI reviewed", output)

    def test_las_expectativas_previas_del_merge_gobernado_siguen_valiendo(self) -> None:
        # Criterio 8: el control nuevo no debilita ninguna expectativa previa: un
        # PR verde sigue fusionando con 0 y un check rojo sigue bloqueando con 1.
        proc = run_gate(self.root, contract(), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_merged(proc, True)
        state = base_state()
        state["check_runs"][SHA] = [{"id": 1, "name": CHECK, "conclusion": "failure", "status": "completed"}]
        proc = run_gate(self.root, contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assert_merged(proc, False)


if __name__ == "__main__":
    unittest.main()
