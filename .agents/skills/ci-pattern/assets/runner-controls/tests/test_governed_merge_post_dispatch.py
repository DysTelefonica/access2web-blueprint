#!/usr/bin/env python3
# ci-pattern asset tests — governed merge post-dispatch validation (Refs #817)
"""#817 (eslabón 2/2 de la cadena): la clave opcional ``post_merge_dispatch`` del
``host-contract.json`` se valida **pre-merge**. Cada nombre debe existir como
workflow en el repo y su YAML debe declarar ``on: workflow_dispatch:``; el
merge se rechaza con exit 1 si falta alguna de las dos. Ausente o vacía es
no-op con un warning a ``stderr``. El disparo post-merge vía ``gh workflow
run`` (slice 3/2 de la cadena, ``Refs #817``; ``Closes #817`` vive en la
punta de la cadena) tiene su propio test file.
"""
import base64
import json
import tempfile
import unittest
from pathlib import Path

import _github_stub
from test_governed_merge import CHECK, SHA, base_state, contract, pull, run_gate


def yaml_workflow(declares: bool) -> str:
    """Returns a minimal Actions YAML either declaring `workflow_dispatch` or only `push`."""
    if declares:
        return ("name: x\n"
                "on:\n"
                "  workflow_dispatch:\n"
                "jobs:\n"
                "  build:\n"
                "    runs-on: ubuntu-latest\n"
                "    steps:\n"
                "      - run: echo ok\n")
    return ("name: x\n"
            "on:\n"
            "  push:\n"
            "    branches: [main]\n"
            "jobs:\n"
            "  build:\n"
            "    runs-on: ubuntu-latest\n"
            "    steps:\n"
            "      - run: echo ok\n")


def with_workflow(name: str, *, declares: bool, present: bool = True) -> dict:
    """Returns the state dict with one workflow registered; the file body is base64."""
    yaml_text = yaml_workflow(declares)
    state = base_state()
    state["workflows"] = ([{"id": 1, "name": name, "path": f".github/workflows/{name}.yml",
                              "state": "active"}]
                          if present else [])
    file_path = f".github/workflows/{name}.yml"
    state["workflow_files"] = {file_path: {"content": base64.b64encode(yaml_text.encode()).decode("ascii"),
                                            "encoding": "base64", "path": file_path,
                                            "name": f"{name}.yml", "type": "file"}}
    return state


def contract_with(post_dispatch: object) -> dict:
    c = contract()
    c["post_merge_dispatch"] = post_dispatch
    return c


BASE_TIP = "0a1b2c3d4e5f60718293a4b5c6d7e8f90a1b2c3d"  # #933: punta de main antes del merge
BASE_BRANCH = pull()["base"]["ref"]  # `ref` del dispatch: la rama base del PR, nunca un SHA


def yaml_with_inputs(keys: list[str]) -> str:
    """#933: YAML con `workflow_dispatch` y las entradas `keys` declaradas bajo `inputs:`."""
    entries = "".join(f"      {key}:\n        required: true\n" for key in keys)
    return ("name: x\n"
            "on:\n"
            "  workflow_dispatch:\n"
            "    inputs:\n"
            f"{entries}"
            "jobs:\n"
            "  build:\n"
            "    runs-on: ubuntu-latest\n"
            "    steps:\n"
            "      - run: echo ok\n")


def with_yaml(name: str, yaml_text: str) -> dict:
    """#933: un workflow del repo con su YAML exacto y la punta de `main` conocida."""
    state = base_state()
    path = f".github/workflows/{name}.yml"
    state["workflows"] = [{"id": 1, "name": name, "path": path, "state": "active"}]
    state["workflow_files"] = {path: {"content": base64.b64encode(yaml_text.encode()).decode("ascii"),
                                       "encoding": "base64", "path": path,
                                       "name": f"{name}.yml", "type": "file"}}
    state["base_tips"] = {"main": BASE_TIP}
    return state


class GovernedMergeDispatchRangeTests(unittest.TestCase):
    """#933: el dispatch lleva `before`/`after` solo a los workflows que los declaran."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def stored(self) -> dict:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))

    def test_a_workflow_declaring_before_and_after_gets_the_range_inputs(self) -> None:
        state = with_yaml("push-compliance", yaml_with_inputs(["before", "after"]))
        proc = run_gate(self.root, contract_with(["push-compliance"]), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertTrue(self.stored()["merges"])
        dispatch = self.stored()["dispatches"][0]
        self.assertEqual(dispatch["inputs"], {"before": BASE_TIP, "after": _github_stub.MERGED_SHA})

    def test_a_workflow_without_range_inputs_is_dispatched_with_no_inputs(self) -> None:
        # Compatibilidad: `deploy` y similares no declaran before/after y no reciben inputs.
        proc = run_gate(self.root, contract_with(["deploy"]),
                        with_yaml("deploy", yaml_with_inputs([])))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIsNone(self.stored()["dispatches"][0]["inputs"])

    def test_the_dispatch_ref_is_the_base_branch_and_after_carries_the_merged_sha(self) -> None:
        # #933: `POST /dispatches` solo admite una rama o una etiqueta en `ref`; con un
        # SHA responde 422 y el detective post-push no llega a ejecutarse. El SHA
        # fusionado viaja solo en `inputs.after`.
        state = with_yaml("push-compliance", yaml_with_inputs(["before", "after"]))
        proc = run_gate(self.root, contract_with(["push-compliance"]), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        dispatch = self.stored()["dispatches"][0]
        self.assertEqual(dispatch["ref"], BASE_BRANCH)
        self.assertNotRegex(dispatch["ref"], r"\A[0-9a-f]{7,40}\Z")
        self.assertEqual(dispatch["inputs"], {"before": BASE_TIP, "after": _github_stub.MERGED_SHA})

    def test_a_workflow_declaring_only_before_blocks_before_the_merge(self) -> None:
        proc = run_gate(self.root, contract_with(["push-compliance"]),
                        with_yaml("push-compliance", yaml_with_inputs(["before"])))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertFalse(self.stored().get("merges", []))
        self.assertIn("after", proc.stdout + proc.stderr)

    def test_an_unreadable_base_tip_fails_closed_before_the_merge(self) -> None:
        state = with_yaml("push-compliance", yaml_with_inputs(["before", "after"]))
        del state["base_tips"]  # la punta de main no se puede leer: no hay rango que enviar
        proc = run_gate(self.root, contract_with(["push-compliance"]), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertFalse(self.stored().get("merges", []))


class GovernedMergePostDispatchTests(unittest.TestCase):
    """#817: workflow_dispatch rebinds the post-merge run set when push is suppressed."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def stored(self) -> dict:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))

    def assert_no_dispatch(self) -> None:
        self.assertEqual(self.stored().get("dispatches", []), [])

    def assert_dispatched_on_base_branch(self, name: str) -> None:
        # #933: el `ref` del cuerpo es la rama base; un SHA ahí es 422 en la API real.
        for dispatch in self.stored().get("dispatches", []):
            if dispatch["name"] == name and dispatch["ref"] == BASE_BRANCH:
                return
        self.fail(f"expected dispatch of {name} on {BASE_BRANCH}; "
                  f"got {self.stored().get('dispatches', [])}")

    def test_post_dispatch_absent_is_a_noop_with_a_warning(self) -> None:
        proc = run_gate(self.root, contract(), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_no_dispatch()
        self.assertIn("post_merge_dispatch", proc.stdout + proc.stderr)

    def test_post_dispatch_empty_list_is_a_noop_with_a_warning(self) -> None:
        proc = run_gate(self.root, contract_with([]), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_no_dispatch()
        self.assertIn("post_merge_dispatch", proc.stdout + proc.stderr)

    def test_post_dispatch_fires_for_each_workflow_after_a_green_merge(self) -> None:
        # Two workflows declared; both declared `workflow_dispatch`; both fire on the base branch.
        state = with_workflow("deploy", declares=True)
        state["workflows"].append({"id": 2, "name": "runner-detective",
                                    "path": ".github/workflows/runner-detective.yml",
                                    "state": "active"})
        state["workflow_files"][".github/workflows/runner-detective.yml"] = {
            "content": base64.b64encode(yaml_workflow(True).encode()).decode("ascii"),
            "encoding": "base64", "path": ".github/workflows/runner-detective.yml",
            "name": "runner-detective.yml", "type": "file"}
        proc = run_gate(self.root, contract_with(["deploy", "runner-detective"]), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assert_dispatched_on_base_branch("deploy")
        self.assert_dispatched_on_base_branch("runner-detective")
        self.assertTrue(self.stored()["merges"])

    def test_post_dispatch_validation_blocks_when_workflow_dispatch_is_missing(self) -> None:
        # The workflow exists on the repo but only declares `push`; the merge
        # must block BEFORE the merge call (no `merges` recorded).
        proc = run_gate(self.root, contract_with(["deploy"]),
                         with_workflow("deploy", declares=False))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertFalse(self.stored().get("merges", []))
        self.assert_no_dispatch()
        self.assertIn("workflow_dispatch", proc.stdout)
        self.assertIn("deploy", proc.stdout)

    def test_post_dispatch_validation_blocks_unknown_workflow(self) -> None:
        state = with_workflow("ci", declares=True)  # the only workflow in the repo
        proc = run_gate(self.root, contract_with(["nonexistent"]), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertFalse(self.stored().get("merges", []))
        self.assert_no_dispatch()
        self.assertIn("nonexistent", proc.stdout + proc.stderr)
        self.assertIn("ci", proc.stdout + proc.stderr)  # the only candidate

    def test_post_dispatch_unique_items_rejects_duplicates_before_the_merge(self) -> None:
        proc = run_gate(self.root, contract_with(["ci", "ci"]),
                         with_workflow("ci", declares=True))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertFalse(self.stored().get("merges", []))
        self.assertIn("duplicate", (proc.stdout + proc.stderr).lower())

    def test_post_dispatch_failure_after_merge_is_exit_4_with_no_rollback(self) -> None:
        # The merge already happened; the dispatch returns 422 -> exit 4.
        state = base_state()
        state.update(with_workflow("deploy", declares=True))
        state["dispatch_error"] = True
        proc = run_gate(self.root, contract_with(["deploy"]), state)
        self.assertEqual(proc.returncode, 4, proc.stdout + proc.stderr)
        # Merge is in state — not rolled back.
        self.assertTrue(self.stored()["merges"])
        self.assert_no_dispatch()
        for expected in ("MERGED", "deploy", "HTTP 422"):
            self.assertIn(expected, proc.stdout + proc.stderr)

    def test_post_dispatch_forbidden_after_merge_is_exit_4_with_the_hint(self) -> None:
        # 403 with the `actions` permission header is the canonical "missing `actions: write`"
        # symptom in the consumer workflow; the message stays actionable.
        state = base_state()
        state.update(with_workflow("deploy", declares=True))
        state["dispatch_forbidden"] = True
        proc = run_gate(self.root, contract_with(["deploy"]), state)
        self.assertEqual(proc.returncode, 4, proc.stdout + proc.stderr)
        self.assertTrue(self.stored()["merges"])
        self.assert_no_dispatch()
        output = proc.stdout + proc.stderr
        self.assertIn("deploy", output)
        self.assertIn("HTTP 403", output)
        self.assertIn("actions=write", output)

    def test_post_dispatch_validation_failure_lets_blocked_findings_keep_finding_first(self) -> None:
        # If the gate already blocked the merge (e.g. red CI), the
        # post-dispatch validation must NOT bypass — exit 1 from the gate
        # wins and no validation happens against the contract.
        red_state = base_state()
        red_state["check_runs"][SHA] = [
            {"id": 1, "name": CHECK, "conclusion": "failure", "status": "completed"}]
        red_state["workflows"] = [{"id": 1, "name": "deploy",
                                    "path": ".github/workflows/deploy.yml", "state": "active"}]
        red_state["workflow_files"] = {".github/workflows/deploy.yml": {
            "content": base64.b64encode(yaml_workflow(False).encode()).decode("ascii"),
            "encoding": "base64", "path": ".github/workflows/deploy.yml",
            "name": "deploy.yml", "type": "file"}}
        proc = run_gate(self.root, contract_with(["deploy"]), red_state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertFalse(self.stored().get("merges", []))
        self.assert_no_dispatch()


if __name__ == "__main__":
    unittest.main()
