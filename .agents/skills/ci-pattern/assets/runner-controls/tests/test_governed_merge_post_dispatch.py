#!/usr/bin/env python3
# ci-pattern asset tests — governed merge post-dispatch validation (Refs #817)
"""#817 (eslabón 2/2 de la cadena): la clave opcional ``post_merge_dispatch`` del
``host-contract.json`` se valida **pre-merge**. Cada nombre debe existir como
workflow en el repo y su YAML debe declarar ``on: workflow_dispatch:``; el
merge se rechaza con exit 1 si falta alguna de las dos. Ausente o vacía es
no-op con un warning a ``stderr``. El disparo post-merge vía ``gh workflow
run`` (slice 3/2 de la cadena, ``Refs #817``; ``Closes #817`` vive en la
punta de la cadena) tiene su propio test file.

#966: la validacion pre-merge se hace contra la forma real que devuelve la
API —un objeto ``{"total_count": N, "workflows": [...]}``, no una lista— y
recorre todas sus paginas (``per_page=100``); sin la clave ``workflows``, o
con un valor que no sea lista, sigue bloqueando nombrando la forma esperada.
"""
import base64
import json
import subprocess
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

    def assert_dispatched_on_base_branch(self, dispatch_id: str) -> None:
        # #1086: el identificador del dispatch es el nombre del fichero con su
        # extension (`ci.yml`); el basename sin extension no lo acepta la API.
        # #933: el `ref` del cuerpo es la rama base; un SHA ahí es 422 en la API real.
        for dispatch in self.stored().get("dispatches", []):
            if dispatch["name"] == dispatch_id and dispatch["ref"] == BASE_BRANCH:
                return
        self.fail(f"expected dispatch of {dispatch_id} on {BASE_BRANCH}; "
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
        self.assert_dispatched_on_base_branch("deploy.yml")
        self.assert_dispatched_on_base_branch("runner-detective.yml")
        self.assertTrue(self.stored()["merges"])

    def test_the_dispatch_carries_the_workflow_file_name_with_its_extension(self) -> None:
        """#1086: la API resuelve el workflow por su nombre de fichero, con extension.

        El endpoint de dispatch acepta el nombre del fichero (``ci.yml``), su
        ruta o su id numerico. Sin la extension responde 404 y el detective
        post-merge no dispara nunca: la fusion queda sin la auditoria que la
        precedia, y el rojo real llega a ``main`` sin verse (#1086).
        """
        proc = run_gate(self.root, contract_with(["deploy"]),
                        with_workflow("deploy", declares=True))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.stored()["dispatches"][0]["name"], "deploy.yml")

    def test_the_dispatch_id_is_the_real_file_name_not_a_guessed_extension(self) -> None:
        """#1086: el identificador sale del listado, no de anadir ``.yml`` a ciegas.

        Un consumer con ``deploy.yaml`` merece su dispatch igual: el nombre del
        fichero se lee del listado de workflows, no se compone.
        """
        state = base_state()
        path = ".github/workflows/deploy.yaml"
        state["workflows"] = [{"id": 1, "name": "deploy", "path": path, "state": "active"}]
        state["workflow_files"] = {path: {
            "content": base64.b64encode(yaml_workflow(True).encode()).decode("ascii"),
            "encoding": "base64", "path": path,
            "name": "deploy.yaml", "type": "file"}}
        state["base_tips"] = {"main": BASE_TIP}
        proc = run_gate(self.root, contract_with(["deploy"]), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.stored()["dispatches"][0]["name"], "deploy.yaml")

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


class GovernedMergeWorkflowsListingTests(unittest.TestCase):
    """#966: el listado de workflows es un objeto y se lee entero, pagina a pagina."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def stored(self) -> dict:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))

    def assert_listing_blocked(self, proc: subprocess.CompletedProcess) -> None:
        output = proc.stdout + proc.stderr
        self.assertEqual(proc.returncode, 1, output)
        self.assertFalse(self.stored().get("merges", []), output)
        self.assertIn("workflows listing is not the API shape", output)
        for shape in ('"total_count"', '"workflows"'):
            self.assertIn(shape, output)  # el mensaje nombra la forma esperada

    def test_the_real_api_listing_shape_lets_the_governed_merge_through(self) -> None:
        # #966: `GET /repos/<repo>/actions/workflows` responde un objeto
        # `{"total_count": N, "workflows": [...]}`; exigir una lista bloqueaba
        # todo consumer con `post_merge_dispatch` declarado en su contrato.
        proc = run_gate(self.root, contract_with(["deploy"]),
                        with_workflow("deploy", declares=True))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertTrue(self.stored()["merges"])

    def test_a_bare_list_listing_blocks_and_names_the_expected_shape(self) -> None:
        # La forma anterior (lista desnuda) no es lo que devuelve la API: sigue
        # bloqueando, y el mensaje dice que se esperaba el objeto (#966).
        state = with_workflow("deploy", declares=True)
        state["workflows_body"] = state["workflows"]
        self.assert_listing_blocked(run_gate(self.root, contract_with(["deploy"]), state))

    def test_a_listing_without_the_workflows_key_blocks(self) -> None:
        state = with_workflow("deploy", declares=True)
        state["workflows_body"] = {"total_count": 1}
        self.assert_listing_blocked(run_gate(self.root, contract_with(["deploy"]), state))

    def test_a_workflows_value_that_is_not_a_list_blocks(self) -> None:
        state = with_workflow("deploy", declares=True)
        state["workflows_body"] = {"total_count": 1,
                                   "workflows": {"path": ".github/workflows/deploy.yml"}}
        self.assert_listing_blocked(run_gate(self.root, contract_with(["deploy"]), state))

    def test_a_workflow_on_the_second_page_is_found(self) -> None:
        # #966: con mas de 100 workflows el listado llega repartido en paginas;
        # el declarado en el contrato esta en la segunda y el merge sale igual.
        # El defecto de la API (30 por pagina) hacia que un repo con 31
        # workflows diera un falso «no existe».
        state = base_state()
        state["workflows"] = [{"id": number, "name": f"wf-{number}",
                               "path": f".github/workflows/wf-{number}.yml", "state": "active"}
                              for number in range(1, 101)]
        state["workflows"].append({"id": 101, "name": "deploy",
                                   "path": ".github/workflows/deploy.yml", "state": "active"})
        path = ".github/workflows/deploy.yml"
        state["workflow_files"] = {path: {"content": base64.b64encode(yaml_workflow(True).encode()).decode("ascii"),
                                           "encoding": "base64", "path": path,
                                           "name": "deploy.yml", "type": "file"}}
        proc = run_gate(self.root, contract_with(["deploy"]), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.stored()["dispatches"][0]["name"], "deploy.yml")
        self.assertEqual(self.stored()["workflows_queries"],
                         [{"per_page": ["100"], "page": ["1"]},
                          {"per_page": ["100"], "page": ["2"]}])


# #1440: el `ci` de cadete saltaba sus jobs en `workflow_dispatch`; la verificacion
# del arbol integrado no corria. `post_merge_verify` la separa de los detectores.
VERIFY_TREE_MOVED = "9" * 40


def job_yaml(condition: str = "") -> str:
    """Un workflow con `workflow_dispatch` y un job `tests`, con su `if:` si se pide."""
    guard = f"    if: {condition}\n" if condition else ""
    return ("name: ci\n"
            "on:\n"
            "  pull_request:\n"
            "  workflow_dispatch:\n"
            "jobs:\n"
            "  tests:\n"
            f"{guard}"
            "    runs-on: ubuntu-latest\n"
            "    steps:\n"
            "      - run: echo ok\n")


def verify_state(ci_yaml: str, *, behind: int, ci_present: bool = True) -> dict:
    """`ci` (verificacion) y `push-compliance` (detector) en el repo; PR atrasado `behind`."""
    state = with_yaml("push-compliance", yaml_with_inputs(["before", "after"]))
    if ci_present:
        path = ".github/workflows/ci.yml"
        state["workflows"].append({"id": 2, "name": "ci", "path": path, "state": "active"})
        state["workflow_files"][path] = {
            "content": base64.b64encode(ci_yaml.encode()).decode("ascii"),
            "encoding": "base64", "path": path, "name": "ci.yml", "type": "file"}
    state["compares"] = {
        f"main...{SHA}": {"behind_by": behind, "files": [{"filename": "src/app.php"}]},
        f"{SHA}...main": {"files": [{"filename": "docs/guide.md"}] if behind else []}}
    if behind:
        state["commits"][_github_stub.MERGED_SHA] = {
            "sha": _github_stub.MERGED_SHA, "commit": {"tree": {"sha": VERIFY_TREE_MOVED}}}
    with_policy(state, {"required_jobs": ["tests"], "events": {}})
    return state


POLICY_PATH = ".github/required-jobs-policy.json"


def with_policy(state: dict, policy: object) -> dict:
    """#1440: la politica de jobs requeridos del repo, leida por la API de contenidos."""
    text = policy if isinstance(policy, str) else json.dumps(policy)
    state.setdefault("workflow_files", {})[POLICY_PATH] = {
        "content": base64.b64encode(text.encode()).decode("ascii"), "encoding": "base64",
        "path": POLICY_PATH, "name": "required-jobs-policy.json", "type": "file"}
    return state


def verify_contract(verify: object = ("ci",), dispatch: object = ("push-compliance",)) -> dict:
    data = contract()
    if dispatch is not None:
        data["post_merge_dispatch"] = list(dispatch)
    data["post_merge_verify"] = list(verify) if isinstance(verify, tuple) else verify
    return data


class GovernedMergePostMergeVerifyTests(unittest.TestCase):
    """#1440: la verificacion del arbol integrado solo se lanza si la PR iba atrasada."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def stored(self) -> dict:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))

    def dispatched(self) -> list[str]:
        return [item["name"] for item in self.stored().get("dispatches", [])]

    def test_verify_is_dispatched_when_the_pr_was_behind(self) -> None:
        proc = run_gate(self.root, verify_contract(), verify_state(job_yaml(), behind=2))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertTrue(self.stored()["merges"])
        ci = [item for item in self.stored()["dispatches"] if item["name"] == "ci.yml"]
        self.assertEqual(len(ci), 1, self.stored()["dispatches"])
        self.assertEqual(ci[0]["ref"], BASE_BRANCH)

    def test_verify_is_not_dispatched_when_the_pr_was_up_to_date(self) -> None:
        proc = run_gate(self.root, verify_contract(), verify_state(job_yaml(), behind=0))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertTrue(self.stored()["merges"])
        self.assertNotIn("ci.yml", self.dispatched())
        self.assertIn("NOT DISPATCHED: ci (post_merge_verify)", proc.stdout)

    def test_post_merge_dispatch_detectors_are_dispatched_after_every_merge(self) -> None:
        for behind in (0, 3):
            with self.subTest(behind=behind):
                proc = run_gate(self.root, verify_contract(),
                                verify_state(job_yaml(), behind=behind))
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
                detector = [item for item in self.stored()["dispatches"]
                            if item["name"] == "push-compliance.yml"]
                self.assertEqual(len(detector), 1, self.stored()["dispatches"])
                self.assertEqual(detector[0]["inputs"],
                                 {"before": BASE_TIP, "after": _github_stub.MERGED_SHA})

    def test_a_verify_job_skipped_on_workflow_dispatch_is_refused_before_the_merge(self) -> None:
        conditions = ("github.event_name != 'workflow_dispatch'",
                      '${{ github.event_name != "workflow_dispatch" }}',
                      "github.event_name!='workflow_dispatch' && always()")
        for condition in conditions:
            with self.subTest(condition=condition):
                proc = run_gate(self.root, verify_contract(),
                                verify_state(job_yaml(condition), behind=2))
                self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
                self.assertFalse(self.stored().get("merges", []))
                self.assertEqual(self.stored().get("dispatches", []), [])
                self.assertIn("post_merge_verify workflow 'ci' skips the required job 'tests' on "
                              "workflow_dispatch: the post-merge verification would not run it "
                              "(#1440)", proc.stdout)

    def test_a_skip_in_a_block_scalar_condition_is_refused(self) -> None:
        yaml_text = job_yaml().replace(
            "    runs-on:",
            "    if: >-\n      github.event_name != 'workflow_dispatch'\n    runs-on:")
        proc = run_gate(self.root, verify_contract(), verify_state(yaml_text, behind=0))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("skips the required job 'tests'", proc.stdout)

    def test_a_non_required_job_skipped_on_dispatch_is_accepted(self) -> None:
        # Verificador #1440, bloqueo 1: solo cuentan los jobs de `required_jobs`.
        yaml_text = job_yaml() + ("  other:\n"
                                  "    if: github.event_name != 'workflow_dispatch'\n"
                                  "    runs-on: ubuntu-latest\n"
                                  "    steps:\n"
                                  "      - run: echo other\n")
        proc = run_gate(self.root, verify_contract(), verify_state(yaml_text, behind=2))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("ci.yml", self.dispatched())

    def test_an_event_filter_that_does_not_admit_dispatch_is_refused(self) -> None:
        # Verificador #1440, bloqueo 2: `== 'pull_request'` tambien salta el dispatch.
        conditions = ("github.event_name == 'pull_request'",
                      "'workflow_dispatch' != github.event_name",
                      "!(github.event_name == 'workflow_dispatch')",
                      "github.event_name == 'push'  # workflow_dispatch")
        for condition in conditions:
            with self.subTest(condition=condition):
                proc = run_gate(self.root, verify_contract(),
                                verify_state(job_yaml(condition), behind=2))
                self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
                self.assertFalse(self.stored().get("merges", []))
                self.assertIn("skips the required job 'tests'", proc.stdout)

    def test_a_filter_that_admits_dispatch_or_no_event_filter_is_accepted(self) -> None:
        conditions = ("github.event_name == 'workflow_dispatch' || github.event_name == 'push'",
                      "always()",
                      "${{ contains('push workflow_dispatch', github.event_name) }}")
        for condition in conditions:
            with self.subTest(condition=condition):
                proc = run_gate(self.root, verify_contract(),
                                verify_state(job_yaml(condition), behind=2))
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_four_space_jobs_and_quoted_names_are_scanned(self) -> None:
        yaml_text = ("on:\n"
                     "    workflow_dispatch:\n"
                     "jobs:\n"
                     "    'tests':\n"
                     "        if: github.event_name != 'workflow_dispatch'\n"
                     "        runs-on: ubuntu-latest\n"
                     "        steps:\n"
                     "            - run: echo ok\n")
        proc = run_gate(self.root, verify_contract(), verify_state(yaml_text, behind=2))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("skips the required job 'tests' on", proc.stdout)

    def test_an_unreadable_required_jobs_policy_fails_closed_before_the_merge(self) -> None:
        cases = {"missing": None, "not json": "{nope", "no required_jobs": {"events": {}}}
        for label, policy in cases.items():
            with self.subTest(case=label):
                state = verify_state(job_yaml(), behind=2)
                if policy is None:
                    del state["workflow_files"][POLICY_PATH]
                else:
                    with_policy(state, policy)
                proc = run_gate(self.root, verify_contract(), state)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertFalse(self.stored().get("merges", []))
                self.assertIn(".github/required-jobs-policy.json", proc.stderr)

    def test_a_step_level_condition_is_not_a_skipped_job(self) -> None:
        yaml_text = job_yaml().replace(
            "      - run: echo ok\n",
            "      - run: echo ok\n        if: github.event_name != 'workflow_dispatch'\n")
        proc = run_gate(self.root, verify_contract(), verify_state(yaml_text, behind=2))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_a_missing_or_undispatchable_verify_workflow_is_refused_before_the_merge(self) -> None:
        cases = {"missing": (verify_state(job_yaml(), behind=2, ci_present=False),
                             "post_merge_verify entry 'ci' is not a workflow in this repo"),
                 "no dispatch": (verify_state(yaml_workflow(False), behind=2),
                                 "workflow 'ci' does not declare 'workflow_dispatch'")}
        for label, (state, expected) in cases.items():
            with self.subTest(case=label):
                proc = run_gate(self.root, verify_contract(), state)
                self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
                self.assertFalse(self.stored().get("merges", []))
                self.assertIn(expected, proc.stdout)

    def test_a_workflow_in_both_lists_is_refused(self) -> None:
        proc = run_gate(self.root, verify_contract(dispatch=("push-compliance", "ci")),
                        verify_state(job_yaml(), behind=2))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("'ci' is in both post_merge_dispatch and post_merge_verify", proc.stdout)

    def test_a_malformed_verify_list_fails_closed_before_any_api_call(self) -> None:
        for value in ("ci", [""], [1]):
            with self.subTest(value=value):
                proc = run_gate(self.root, verify_contract(verify=value),
                                verify_state(job_yaml(), behind=2))
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertIn("post_merge_verify must be a list of non-empty strings",
                              proc.stderr)
                self.assertEqual(self.stored()["calls"], [])

    def test_a_verify_dispatch_error_after_the_merge_is_exit_4(self) -> None:
        state = verify_state(job_yaml(), behind=2)
        state["dispatch_error"] = True
        proc = run_gate(self.root, verify_contract(), state)
        self.assertEqual(proc.returncode, 4, proc.stdout + proc.stderr)
        self.assertTrue(self.stored()["merges"])
        self.assertIn("ci", proc.stderr)


if __name__ == "__main__":
    unittest.main()
