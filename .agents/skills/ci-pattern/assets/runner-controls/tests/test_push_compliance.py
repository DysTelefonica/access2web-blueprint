#!/usr/bin/env python3
# ci-pattern asset tests — post-push compliance detective (DysTelefonica/team-skills#303)
"""Post-push detective suite (black-box): a violating push opens an issue
with SHA, author and check evidence (exit 1); a broken environment fails
closed (exit 2). The API is a local HTTP stub with the real shapes; stdlib only."""
import http.client
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from _github_stub import github_stub, stub_env

ASSET = Path(__file__).resolve().parent.parent / "check_push_compliance.py"
REPO = "DysTelefonica/team-skills"
SHA = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
BASE = "0a1b2c3d4e5f60718293a4b5c6d7e8f90a1b2c3d"
CHECK = "unit suite (≤3min)"
# #909: commit de merge y mensajes reales de cadete (#536 gobernado, #373 web).
MERGE_SHA = "5e6f7a8b9c0d1e2f3a4b5c6d7e8f9a0b1c2d3e4f"
GOVERNED_MESSAGE = ("Merge PR #536 (governed merge)\n\n"
                    "ci(host-contract): docs-gates y required en required_checks y P28")
WEB_MESSAGE = ("Merge pull request #373 from DysTelefonica/chore/366-adopt-issue-spec-canon-config\n\n"
               "ci(issue-spec): require canonical sections and document chain context "
               "(Closes #366, chain 4/4)")

# #1124: los tres commits reales de cadete#599 -> cadete#629. Los shas completos
# no hacen falta para el contrato; el prefijo publicado permite trazar el caso.
CHAIN_FIRST_SHA = "4fbba72e" + "0" * 32
CHAIN_SECOND_SHA = "f99ca724" + "0" * 32
CHAIN_HEAD_SHA = "55a4cb40" + "0" * 32
CHAIN_MERGE_SHA = "66a4cb40" + "0" * 32
DIRECT_SHA = "77a4cb40" + "0" * 32

def write_fixture(root: Path, event: dict, contract: dict) -> None:
    (root / "event.json").write_text(json.dumps(event), encoding="utf-8")
    (root / "contract.json").write_text(json.dumps(contract), encoding="utf-8")


def push_event(ref: str = "refs/heads/main", commits: list | None = None) -> dict:
    commit = {"id": SHA, "message": "feat: change", "author": {"name": "Dev A", "email": "dev@example.com"},
              "url": f"https://github.com/{REPO}/commit/{SHA}", "distinct": True}
    return {"ref": ref, "before": BASE, "after": SHA, "created": False, "forced": False, "deleted": False,
            "commits": commits if commits is not None else [commit], "head_commit": commit,
            "pusher": {"name": "deva", "email": "dev@example.com"},
            "repository": {"full_name": REPO, "default_branch": "main"}}


def contract() -> dict:
    return {"contract_version": 1, "protected_branches": ["main"],
            "required_checks": [{"name": CHECK, "class": "documented-only"}]}


# #928: forma REAL de GET repos/{r}/commits/{sha}: el arbol vive en
# commit.tree.sha (top-level sin campo `tree`). Shas de arbol de 40 hex.
TREE_PR = "a1b2c3d4e5f60718293a4b5c6d7e8f9012345678"
TREE_PUSH = "0f1e2d3c4b5a69788796a5b4c3d2e1f0a9b8c7d6"


def commit_obj(sha: str, tree: str) -> dict:
    return {"sha": sha, "commit": {"tree": {"sha": tree}}}


def base_state() -> dict:
    return {"pulls": {SHA: [{"number": 7, "merged": True, "merged_at": "2026-10-06T10:00:00Z",
                             "user": {"login": "ardelperal"}, "base": {"ref": "main"}}]},
            "pr_details": {"7": {"number": 7, "merged": True, "user": {"login": "ardelperal"},
                                 "head": {"sha": SHA}, "merge_commit_sha": SHA}},
            "check_runs": {SHA: [{"id": 1, "name": CHECK, "conclusion": "success", "status": "completed"}]},
            "commits": {SHA: commit_obj(SHA, TREE_PR)},
            "created_issues": [], "calls": [], "merges": [],
            "compare": {"total_commits": 1,
                        "commits": [{"sha": SHA,
                                         "commit": {"author": {"name": "Dev A"}}}]}}


def governed_contract(actors: list | None = None) -> dict:
    """Contrato con la clave opt-in `governed_path` activada (#909)."""
    return {**contract(), "governed_path": {
        "enabled": True,
        "actors": actors if actors is not None else ["github-actions[bot]"]}}


def merge_state(*, author: str, merged_by: str, message: str) -> dict:
    """Estado de un push cuyo PR se fusiono con el commit de merge `MERGE_SHA`."""
    state = base_state()
    state["pulls"][SHA] = [{"number": 7, "merged": True, "merged_at": "2026-10-19T22:21:58Z",
                            "user": {"login": author}, "base": {"ref": "main"}}]
    state["pr_details"]["7"] = {"number": 7, "merged": True, "user": {"login": author},
                                 "merged_by": {"login": merged_by},
                                 "head": {"sha": SHA}, "merge_commit_sha": MERGE_SHA}
    state["commits"][MERGE_SHA] = {"sha": MERGE_SHA, "commit": {"message": message}}
    return state


def run_gate(root: Path, event: dict, contract_data: dict, state: dict,
             *extra: str) -> subprocess.CompletedProcess:
    write_fixture(root, event, contract_data)
    with github_stub(root / "state.json", state) as base:
        return subprocess.run([sys.executable, str(ASSET), "--event", str(root / "event.json"),
                               "--contract", str(root / "contract.json"), *extra],
                              capture_output=True, text=True, env=stub_env(base), timeout=60)


def settle(*extra: str) -> tuple:
    """Flags de la ventana de settle, en segundos de test (los reales son minutos).

    El poll va siempre: en tests reales es de minutos y sin el se bloquearia.
    """
    flags = list(extra) or ["--wait-seconds", "0.3"]
    if "--poll-seconds" not in flags:
        flags += ["--poll-seconds", "0.05"]
    return tuple(flags)


class MeasuredRaceTests(unittest.TestCase):
    """#1019: los cinco pushes que produjeron incidente medidos en APAP_WEB.

    Datos grabados de ardelperal/APAP_WEB (incidentes #2243, #2247, #2249,
    #2251 y #2280; 2026-10-21/22). Se fijan aqui como datos: ningun test lee
    historia de git ni consulta la API. Cuatro de los cinco mergearon segundos
    despues del push del head y el rollup quedo verde minutos mas tarde; el
    quinto lo mergearon con el rollup ya cancelado, que es un hecho consumado.
    """

    # (incidente, PR, head, conclusion del rollup despues del merge, veredicto esperado)
    CASES = [
        (2243, 2224, "619a8d57606e53ff796f45259c999a9246fc1955", "success", 0),
        (2247, 2234, "529968ba9f8870a4d1a247f2918e356cca241ca5", "success", 0),
        (2251, 2248, "9bd7091d64a0a0c4de25f4fed85dc89cd7e5204d", "success", 0),
        (2280, 2264, "9e05bd7065ca63b4d29e63d717e1b7de7583eb05", "success", 0),
        (2249, 2246, "06daebd01a0a25324b8e66c3bc565d7b53e07d9f", "cancelled", 1),
    ]

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def test_measured_pushes_keep_their_measured_verdict(self) -> None:
        for incident, pr, head, conclusion, expected in self.CASES:
            with self.subTest(incident=f"#{incident} PR#{pr}"):
                state = base_state()
                sha = head
                state["pulls"][sha] = [{"number": pr, "merged": True,
                                        "merged_at": "2026-10-22T14:52:01Z",
                                        "user": {"login": "ardelperal"}, "base": {"ref": "main"}}]
                state["pr_details"][str(pr)] = {"number": pr, "merged": True,
                                                "user": {"login": "ardelperal"},
                                                "merged_by": {"login": "ardelperal"},
                                                "head": {"sha": sha}, "merge_commit_sha": sha}
                run = {"id": 7, "name": CHECK, "conclusion": conclusion, "status": "completed",
                       "html_url": f"https://github.com/{REPO}/actions/runs/{incident}"}
                published = [] if conclusion == "success" else [run]
                state["commits"][sha] = commit_obj(sha, TREE_PR)
                state["check_runs_script"] = {sha: [published, published, [run]]}
                state["compare"] = {"total_commits": 1,
                                    "commits": [{"sha": sha, "commit": {"author": {"name": "ardelperal"}}}]}
                event = push_event()
                event["before"], event["after"] = BASE, sha
                proc = run_gate(self.root, event, contract(), state, *settle("--wait-seconds", "30"))
                self.assertEqual(proc.returncode, expected, proc.stdout + proc.stderr)
                issues = json.loads((self.root / "state.json").read_text(
                    encoding="utf-8"))["created_issues"]
                if expected == 0:
                    self.assertEqual(issues, [], f"#{incident} no debe abrir incidente")
                else:
                    self.assertEqual(len(issues), 1)
                    self.assertIn(conclusion, issues[0]["body"])
                    self.assertIn(f"runs/{incident}", issues[0]["body"])


class DetectiveTests(unittest.TestCase):
    """#303 control 1: violaciones abren issue de incidente; conforme sale limpio."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def created_issue(self) -> dict:
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(len(state["created_issues"]), 1, state)
        return state["created_issues"][0]

    def test_compliant_and_unprotected_pushes_exit_zero(self) -> None:
        proc = run_gate(self.root, push_event(), contract(), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])
        proc = run_gate(self.root, push_event(ref="refs/heads/topic/x"), contract(), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["calls"], [])

    def test_tracker_contract_rejects_arrival_from_another_branch(self) -> None:
        """#333 (regla 7): con `tracker_branch` declarado, un PR mergeado cuya
        cabeza no es esa rama es violación y abre incidente."""
        contract_data = {**contract(), "tracker_branch": "feat/303-tracker"}
        state = base_state()
        state["pulls"][SHA] = [{"number": 7, "merged": True, "merged_at": "2026-10-06T10:00:00Z",
                                "user": {"login": "ardelperal"}, "base": {"ref": "main"},
                                "head": {"ref": "otra/rama"}}]
        proc = run_gate(self.root, push_event(), contract_data, state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("tracker_branch", self.created_issue()["body"])

    def test_tracker_contract_accepts_arrival_from_the_tracker(self) -> None:
        """Sin violación, el control no cambia de veredicto: conforme sale 0."""
        contract_data = {**contract(), "tracker_branch": "feat/303-tracker"}
        state = base_state()
        state["pulls"][SHA] = [{"number": 7, "merged": True, "merged_at": "2026-10-06T10:00:00Z",
                                "user": {"login": "ardelperal"}, "base": {"ref": "main"},
                                "head": {"ref": "feat/303-tracker"}}]
        proc = run_gate(self.root, push_event(), contract_data, state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_contract_with_an_unusable_tracker_branch_fails_closed(self) -> None:
        """Un `tracker_branch` declarado y vacío es contrato inservible (HR-3)."""
        proc = run_gate(self.root, push_event(), {**contract(), "tracker_branch": "  "}, base_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("tracker_branch", proc.stdout + proc.stderr)

    def test_first_push_of_a_new_unprotected_branch_is_skipped(self) -> None:
        """#361: un push que CREA una rama no protegida no tiene rango que auditar
        (`before` de ceros) y la rama no es gobernada: sale 0 sin llamar a la API."""
        event = push_event(ref="refs/heads/topic/nueva")
        event["before"] = "0" * 40
        event["created"] = True
        proc = run_gate(self.root, event, contract(), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("not protected", proc.stdout)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["calls"], [])

    def test_first_push_of_a_new_protected_branch_still_fails_closed(self) -> None:
        """La duda no es un pase (HR-3): una rama PROTEGIDA creada de cero no tiene
        rango auditable y sigue saliendo 2."""
        event = push_event()
        event["before"] = "0" * 40
        event["created"] = True
        proc = run_gate(self.root, event, contract(), base_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("not auditable", proc.stdout + proc.stderr)

    def test_commit_pushed_directly_to_main_still_opens_incident(self) -> None:
        state = base_state()
        state["pulls"][DIRECT_SHA] = []
        state["compare"] = {"total_commits": 1,
                            "commits": [{"sha": DIRECT_SHA,
                                         "commit": {"author": {"name": "Dev A"}}}]}
        event = push_event(commits=[{"id": DIRECT_SHA, "message": "feat: direct push",
                                     "author": {"name": "Dev A"}, "distinct": True}])
        event["after"] = DIRECT_SHA
        event["head_commit"] = event["commits"][-1]
        proc = run_gate(self.root, event, contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(proc.stdout, (
            f"  VIOLATION  commit {DIRECT_SHA} — author: Dev A\n"
            "INCIDENT ISSUE: https://github.com/DysTelefonica/team-skills/issues/42\n"))
        self.assertEqual(proc.stderr, "")
        issue = self.created_issue()
        self.assertIn(DIRECT_SHA, issue["body"])
        self.assertIn("Dev A", issue["body"])
        self.assertIn("no merged PR associated", issue["body"])
        self.assertIn("PR", issue["title"] + issue["body"])

    def test_merged_pr_with_red_check_opens_incident(self) -> None:
        state = base_state()
        state["check_runs"][SHA] = [{"id": 1, "name": CHECK, "conclusion": "failure", "status": "completed"}]
        proc = run_gate(self.root, push_event(), contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        issue = self.created_issue()
        self.assertIn(CHECK, issue["body"])
        self.assertIn("failure", issue["body"])
        self.assertIn(SHA, issue["body"])

    def test_required_check_that_appears_later_is_waited_for(self) -> None:
        # #1019: el merge se lanzo antes de que el CI publicara el check; esperar, no acusar.
        state = base_state()
        green = [{"id": 7, "name": CHECK, "conclusion": "success", "status": "completed"}]
        pending = [{"id": 7, "name": CHECK, "conclusion": None, "status": "in_progress"}]
        state["check_runs_script"] = {SHA: [[], pending, green]}
        proc = run_gate(self.root, push_event(), contract(), state, *settle("--wait-seconds", "30"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])
        self.assertGreater(sum("/check-runs" in call for call in state["calls"]), 1)

    def test_required_check_never_published_fails_closed_without_accusing(self) -> None:
        # #1019: el techo agotado con el check sin publicar es duda (HR-3), no una
        # infraccion: exit 2, sin issue de incidente y sin acusar a nadie.
        state = base_state()
        state["check_runs"][SHA] = []
        proc = run_gate(self.root, push_event(), contract(), state, *settle())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail-closed", proc.stderr)
        self.assertIn(CHECK, proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_red_check_does_not_wait_and_is_reported_with_its_run(self) -> None:
        # #1019: una conclusion terminal se juzga de una vez y con su evidencia.
        state = base_state()
        state["check_runs"][SHA] = [{"id": 7, "name": CHECK, "conclusion": "failure",
                                     "status": "completed",
                                     "html_url": "https://github.com/DysTelefonica/team-skills/runs/42"}]
        proc = run_gate(self.root, push_event(), contract(), state, *settle("--wait-seconds", "30"))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        issue = self.created_issue()
        self.assertIn("failure", issue["body"])
        self.assertIn("runs/42", issue["body"])
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(sum("/check-runs" in call for call in state["calls"]), 1)

    def test_negative_wait_is_rejected(self) -> None:
        proc = run_gate(self.root, push_event(), contract(), base_state(),
                        "--wait-seconds", "-1", "--poll-seconds", "0.05")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail-closed", proc.stderr)

    def test_pr_merged_into_another_branch_is_a_violation(self) -> None:
        # #326 bypass: un PR mergeado en OTRA rama no protege un push directo.
        state = base_state()
        state["pulls"][SHA] = [{"number": 7, "merged": True, "merged_at": "2026-10-06T10:00:00Z",
                                "user": {"login": "ardelperal"}, "base": {"ref": "fix/other"}}]
        proc = run_gate(self.root, push_event(), contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        issue = self.created_issue()
        self.assertIn("no merged PR into main", issue["body"])

    def test_truncated_push_payload_fails_closed(self) -> None:
        # #326 truncado: el payload trae 20 commits (el maximo de GitHub)
        # pero el compare reporta 21: rango no auditable, exit 2.
        shas = [f"{i:040x}" for i in range(1, 21)]
        commits = [{"id": s, "message": "x", "author": {"name": "Dev", "email": "d@e"}} for s in shas]
        event = push_event(commits=commits)
        event["head_commit"] = commits[-1]
        event["after"] = shas[-1]
        state = base_state()
        state["compare"] = {"total_commits": 21, "commits": []}
        proc = run_gate(self.root, event, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail", (proc.stdout + proc.stderr).lower())

    def test_multi_page_pulls_response_is_fully_audited(self) -> None:
        # #326 menores: --pagina concatena documentos JSON; el PR mergeado
        # en main esta en la segunda pagina y debe encontrarse.
        state = base_state()
        del state["pulls"]
        state["pulls_pages"] = {SHA: [[], [{"number": 7, "merged": True,
                                            "merged_at": "2026-10-06T10:00:00Z",
                                            "user": {"login": "ardelperal"}, "base": {"ref": "main"}}]]}
        proc = run_gate(self.root, push_event(), contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_stub_link_header_ignores_the_request_host_header(self) -> None:
        # CodeQL py/http-response-splitting: el stub no debe copiar el Host de
        # la peticion a la cabecera Link; la URL sale de la direccion del servidor.
        state = base_state()
        del state["pulls"]
        state["pulls_pages"] = {SHA: [[], []]}
        with tempfile.TemporaryDirectory() as tmp, \
                github_stub(Path(tmp) / "state.json", state) as api:
            netloc = api.split("//", 1)[1]
            conn = http.client.HTTPConnection(netloc)
            conn.putrequest("GET", f"/repos/{REPO}/commits/{SHA}/pulls", skip_host=True)
            conn.putheader("Host", "evil.example")
            conn.endheaders()
            link = conn.getresponse().getheader("Link")
            conn.close()
        self.assertIn(f"<http://{netloc}/", link)
        self.assertNotIn("evil.example", link)

    def test_range_of_25_commits_is_audited_fully(self) -> None:
        # El payload de un push trae 20 como maximo; el rango completo se
        # audita con la lista commits del compare (25 aqui) y sale limpio.
        shas = [f"{i:040x}" for i in range(1, 26)]
        event = push_event(commits=[{"id": s, "message": "x",
                                     "author": {"name": "Dev", "email": "d@e"}}
                                    for s in shas[:20]])
        event["head_commit"] = event["commits"][-1]
        event["after"] = shas[-1]
        state = base_state()
        pr = {"number": 7, "merged": True, "merged_at": "2026-10-06T10:00:00Z",
              "user": {"login": "ardelperal"}, "base": {"ref": "main"}}
        state["pulls"] = {s: [pr] for s in shas}
        # #928: el chequeo de arbol ya mide: cada commit del rango necesita su
        # arbol (aqui el de la PR, rango limpio).
        state["commits"].update({s: commit_obj(s, TREE_PR) for s in shas})
        state["compare"] = {"total_commits": 25,
                            "commits": [{"sha": s, "commit": {"author": {"name": "Dev A"}}}
                                        for s in shas]}
        proc = run_gate(self.root, event, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_range_beyond_the_250_cap_fails_closed(self) -> None:
        # compare reporta 300 pero devuelve 250 (tope de la API): el rango
        # no se puede auditar entero -> exit 2, no un verde a medias.
        shas = [f"{i:040x}" for i in range(1, 301)]
        event = push_event(commits=[{"id": s, "message": "x",
                                     "author": {"name": "Dev", "email": "d@e"}}
                                    for s in shas])
        event["head_commit"] = event["commits"][-1]
        event["after"] = shas[-1]
        state = base_state()
        state["compare"] = {"total_commits": 300,
                            "commits": [{"sha": s, "commit": {"author": {"name": "Dev A"}}}
                                        for s in shas[:250]]}
        proc = run_gate(self.root, event, contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail", (proc.stdout + proc.stderr).lower())

    def test_fail_closed_environment(self) -> None:
        # Evento malformado, cuerpo de API no JSON, error de API y fallo al
        # abrir la issue: duda o fallo de entorno, nunca verde (HR-3).
        bad_event = push_event()
        del bad_event["ref"]
        proc = run_gate(self.root, bad_event, contract(), base_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        state = base_state()
        state["invalid_body"] = True
        proc = run_gate(self.root, push_event(), contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail", (proc.stdout + proc.stderr).lower())
        state = base_state()
        state["api_error"] = True
        proc = run_gate(self.root, push_event(), contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        # Violacion detectada pero la issue no abre: fail closed.
        state = base_state()
        state["pulls"][SHA] = []
        state["issue_post_error"] = True
        proc = run_gate(self.root, push_event(), contract(), state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail", (proc.stdout + proc.stderr).lower())


class MergedChainCoverageTests(unittest.TestCase):
    """#1124: una cadena fusionada en orden cubre todos sus commits en main."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def chain_state(self) -> dict:
        first_pull = {
            "number": 599, "merged": True, "merged_at": "2026-10-23T18:00:00Z",
            "user": {"login": "dys-agentes"},
            "head": {"ref": "skill-fleet/cadete-vendored", "sha": CHAIN_HEAD_SHA},
            "base": {"ref": "skill-fleet/cadete", "sha": CHAIN_FIRST_SHA},
            "merge_commit_sha": CHAIN_SECOND_SHA,
        }
        final_pull = {
            "number": 629, "merged": True, "merged_at": "2026-10-23T19:00:00Z",
            "user": {"login": "dys-agentes"},
            "head": {"ref": "skill-fleet/cadete", "sha": CHAIN_HEAD_SHA},
            "base": {"ref": "main", "sha": BASE},
            "merge_commit_sha": CHAIN_MERGE_SHA,
        }
        state = base_state()
        state["pulls"] = {
            CHAIN_FIRST_SHA: [first_pull],
            CHAIN_SECOND_SHA: [first_pull],
            CHAIN_HEAD_SHA: [first_pull],
            CHAIN_MERGE_SHA: [final_pull],
        }
        state["pr_details"] = {
            "599": first_pull,
            "629": final_pull,
        }
        state["check_runs"] = {
            CHAIN_HEAD_SHA: [{"id": 1, "name": CHECK,
                               "conclusion": "success", "status": "completed"}],
        }
        state["commits"] = {
            CHAIN_FIRST_SHA: commit_obj(CHAIN_FIRST_SHA, TREE_PR),
            CHAIN_SECOND_SHA: commit_obj(CHAIN_SECOND_SHA, TREE_PR),
            CHAIN_HEAD_SHA: commit_obj(CHAIN_HEAD_SHA, TREE_PR),
            CHAIN_MERGE_SHA: {
                "sha": CHAIN_MERGE_SHA,
                "parents": [{"sha": BASE}, {"sha": CHAIN_HEAD_SHA}],
                "commit": {"tree": {"sha": TREE_PR},
                           "message": "Merge pull request #629 from skill-fleet/cadete"},
            },
        }
        compare = {
            "total_commits": 4,
            "commits": [{"sha": sha, "commit": {"author": {"name": "Dev A"}}}
                        for sha in (CHAIN_FIRST_SHA, CHAIN_SECOND_SHA,
                                    CHAIN_HEAD_SHA, CHAIN_MERGE_SHA)],
        }
        head_compare = {
            "status": "ahead", "behind_by": 0, "total_commits": 3,
            "commits": [{"sha": sha, "commit": {"author": {"name": "Dev A"}}}
                        for sha in (CHAIN_FIRST_SHA, CHAIN_SECOND_SHA, CHAIN_HEAD_SHA)],
        }
        state["compare"] = compare
        state["compares"] = {
            f"{BASE}...{CHAIN_HEAD_SHA}": head_compare,
            f"{BASE}...{CHAIN_MERGE_SHA}": compare,
        }
        return state

    def chain_event(self) -> dict:
        shas = (CHAIN_FIRST_SHA, CHAIN_SECOND_SHA, CHAIN_HEAD_SHA, CHAIN_MERGE_SHA)
        commits = [{"id": sha, "message": "feat: chain commit",
                    "author": {"name": "Dev A"}, "distinct": True} for sha in shas]
        event = push_event(commits=commits)
        event["after"] = CHAIN_MERGE_SHA
        event["head_commit"] = commits[-1]
        return event

    def test_two_segment_chain_merged_in_order_does_not_open_incident(self) -> None:
        proc = run_gate(self.root, self.chain_event(), contract(), self.chain_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(proc.stdout,
                         "OK: 4 pushed commit(s) on 'main' carry a merged PR and green checks\n")
        self.assertEqual(proc.stderr, "")
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])
        self.assertEqual(state["merges"], [])

    def test_merged_pr_head_covers_its_chain_range(self) -> None:
        state = self.chain_state()
        state["compare"] = {
            "total_commits": 3,
            "commits": [{"sha": sha, "commit": {"author": {"name": "Dev A"}}}
                        for sha in (CHAIN_FIRST_SHA, CHAIN_SECOND_SHA, CHAIN_HEAD_SHA)],
        }
        state["merged_pulls"] = [state["pr_details"]["629"]]
        shas = (CHAIN_FIRST_SHA, CHAIN_SECOND_SHA, CHAIN_HEAD_SHA)
        commits = [{"id": sha, "message": "feat: chain commit",
                    "author": {"name": "Dev A"}, "distinct": True} for sha in shas]
        event = push_event(commits=commits)
        event["after"] = CHAIN_HEAD_SHA
        event["head_commit"] = commits[-1]
        proc = run_gate(self.root, event, contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(proc.stdout,
                         "OK: 3 pushed commit(s) on 'main' carry a merged PR and green checks\n")
        self.assertEqual(proc.stderr, "")
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])


# #928 Fase 2: invariante de integridad de arbol. Solo el commit de merge
# (merge_commit_sha) con base ancestro del head (compare base...head ahead o
# identical) debe tener el arbol del head. Fixtures con la forma documentada.
C1 = "1111111111111111111111111111111111111111"
C2 = "2222222222222222222222222222222222222222"
BASE_ADVANCED = "3333333333333333333333333333333333333333"
REBASED_PREV = "4444444444444444444444444444444444444444"
TREE_INTERMEDIATE = "5555555555555555555555555555555555555555"
TREE_MAIN = "6666666666666666666666666666666666666666"


def tree_state(*, range_shas: list, merge_parents: list, merge_tree: str, compare_status: str) -> dict:
    """Estado de un push de una PR (#7, head SHA) fusionada con MERGE_SHA.
    `merge_parents[0]` es la base en el momento del merge."""
    state = base_state()
    entry = {"number": 7, "merged": True, "merged_at": "2026-10-19T22:21:58Z",
             "user": {"login": "ardelperal"}, "base": {"ref": "main"},
             "merge_commit_sha": MERGE_SHA}
    state["pulls"] = {s: [entry] for s in range_shas}
    state["pr_details"]["7"] = {"number": 7, "merged": True, "user": {"login": "ardelperal"},
                                "merged_by": {"login": "ardelperal"},
                                "head": {"sha": SHA}, "merge_commit_sha": MERGE_SHA}
    state["commits"][SHA] = commit_obj(SHA, TREE_PR)
    state["commits"][MERGE_SHA] = {"sha": MERGE_SHA, "parents": [{"sha": p} for p in merge_parents],
                                   "commit": {"tree": {"sha": merge_tree}, "message": "Merge pull request #7"}}
    for s in range_shas:
        if s != MERGE_SHA:
            state["commits"][s] = commit_obj(s, TREE_INTERMEDIATE)
    state["compares"] = {f"{merge_parents[0]}...{SHA}": {"status": compare_status, "behind_by": 0}}
    state["compare"] = {"total_commits": len(range_shas), "behind_by": 0,
                        "commits": [{"sha": s, "commit": {"author": {"name": "Dev A"}}} for s in range_shas]}
    return state


def tree_push(range_shas: list) -> dict:
    event = push_event(commits=[{"id": s, "message": "feat: change", "author": {"name": "Dev A"},
                                 "distinct": True} for s in range_shas])
    event["after"] = range_shas[-1]
    event["head_commit"] = event["commits"][-1]
    return event


class TreeIntegrityTests(unittest.TestCase):
    """#928: el chequeo de arbol solo mide el commit de merge con base ancestro
    del head. Intermedios, main avanzado y rebase-merge no se miden (ver docstring
    de check_tree_integrity; el rebase-merge es limitacion conocida)."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def issues(self) -> list:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))["created_issues"]

    def test_conforming_merge_of_up_to_date_branch_exits_zero(self) -> None:
        state = tree_state(range_shas=[MERGE_SHA], merge_parents=[BASE, SHA],
                           merge_tree=TREE_PR, compare_status="ahead")
        proc = run_gate(self.root, tree_push([MERGE_SHA]), contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.issues(), [])

    def test_a_intermediate_commits_of_a_multi_commit_pr_do_not_open_incident(self) -> None:
        # (a) C1 y C2 tienen arbol intermedio legitimo; solo el merge se mide.
        state = tree_state(range_shas=[C1, C2, MERGE_SHA], merge_parents=[BASE, SHA],
                           merge_tree=TREE_PR, compare_status="ahead")
        proc = run_gate(self.root, tree_push([C1, C2, MERGE_SHA]), contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.issues(), [])

    def test_b_merge_after_main_advanced_does_not_open_incident(self) -> None:
        # (b) main avanzo: la base del merge no es ancestro del head (diverged).
        # El arbol del merge difiere legitimamente del head.
        state = tree_state(range_shas=[MERGE_SHA], merge_parents=[BASE_ADVANCED, SHA],
                           merge_tree=TREE_MAIN, compare_status="diverged")
        proc = run_gate(self.root, tree_push([MERGE_SHA]), contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.issues(), [])

    def test_real_drift_on_up_to_date_merge_opens_incident(self) -> None:
        # Positivo: base ancestro del head y el merge trae contenido distinto
        # del revieweado. Debe abrir incidente.
        state = tree_state(range_shas=[MERGE_SHA], merge_parents=[BASE, SHA],
                           merge_tree=TREE_PUSH, compare_status="ahead")
        proc = run_gate(self.root, tree_push([MERGE_SHA]), contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(len(self.issues()), 1, proc.stdout + proc.stderr)
        self.assertIn("difieren en el arbol", self.issues()[0]["body"])

    def test_drift_after_an_intermediate_commit_still_opens_incident(self) -> None:
        # El commit intermedio consume el PR (seen_prs); el merge con drift
        # debe seguir abriendo incidente. Protege el cambio de seen_prs.
        state = tree_state(range_shas=[C1, MERGE_SHA], merge_parents=[BASE, SHA],
                           merge_tree=TREE_PUSH, compare_status="ahead")
        proc = run_gate(self.root, tree_push([C1, MERGE_SHA]), contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(len(self.issues()), 1, proc.stdout + proc.stderr)
        self.assertIn("difieren en el arbol", self.issues()[0]["body"])

    def test_limitation_rebase_merge_drift_is_not_detected(self) -> None:
        # LIMITACION CONOCIDA (#928): en rebase-merge merge_commit_sha es el
        # ultimo commit rebasado; su primer padre es el rebasado anterior, no
        # la base. compare sale diverged y el chequeo no aplica, aunque haya
        # drift. Este test fija ese punto ciego: no abre incidente.
        state = tree_state(range_shas=[MERGE_SHA], merge_parents=[REBASED_PREV],
                           merge_tree=TREE_PUSH, compare_status="diverged")
        proc = run_gate(self.root, tree_push([MERGE_SHA]), contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.issues(), [])


class GovernedPathTests(unittest.TestCase):
    """#909: con la clave opt-in `governed_path`, el detective exige que el merge
    haya llegado por el camino gobernado (actor en la lista o firma en el
    mensaje del commit de merge). Los fixtures son los casos reales de cadete."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def created_issue(self) -> dict:
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(len(state["created_issues"]), 1, state)
        return state["created_issues"][0]

    def test_governed_merge_by_the_bot_passes(self) -> None:
        """Caso real de cadete #536: `github-actions[bot]` fusiona y el commit de
        merge dice «Merge PR #536 (governed merge)»: el camino gobernado se cumple
        y el detective sale limpio."""
        state = merge_state(author="ardelperal", merged_by="github-actions[bot]",
                            message=GOVERNED_MESSAGE)
        proc = run_gate(self.root, push_event(), governed_contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_web_merge_is_a_violation(self) -> None:
        """Caso real de cadete #373: mergeado por la interfaz web con «Merge pull
        request #373 from …» es una violation que abre incidente."""
        state = merge_state(author="ardelperal", merged_by="ardelperal", message=WEB_MESSAGE)
        proc = run_gate(self.root, push_event(), governed_contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        issue = self.created_issue()
        self.assertIn("outside the governed path", issue["body"])
        self.assertIn("ardelperal", issue["body"])
        self.assertIn("Merge pull request #373", issue["body"])

    def test_the_merger_is_not_the_pr_author(self) -> None:
        """El veredicto lee `merged_by` del detalle del PR, no `user.login` de la
        lista de PRs de un commit (allí va el AUTOR): un PR cuyo autor esta en la
        lista de actores y cuyo fusionador no, es violation."""
        state = merge_state(author="ardelperal", merged_by="otro-actor", message=WEB_MESSAGE)
        proc = run_gate(self.root, push_event(), governed_contract(["ardelperal"]), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        issue = self.created_issue()
        self.assertIn("outside the governed path", issue["body"])
        self.assertIn("otro-actor", issue["body"])

    def test_the_merge_marker_is_an_alternative_route(self) -> None:
        """Un fusionador que no esta en `actors` cumple si el commit de merge
        habla el idioma gobernado (la firma que emite `check_governed_merge.py`)."""
        state = merge_state(author="ardelperal", merged_by="ardelperal",
                            message=GOVERNED_MESSAGE)
        proc = run_gate(self.root, push_event(), governed_contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_the_merge_marker_route_is_forgeable_and_that_is_a_known_limit(self) -> None:
        """LIMITE CONOCIDO (#909): la ruta del mensaje es la DEBIL. Quien mergea
        por la web puede editar el titulo del commit de merge y escribir la firma
        gobernada, y el detective lo acepta: este test fija ese limite para que
        no se lea como un fallo. El control detecta un bypass ACCIDENTAL, no uno
        deliberado; la ruta fuerte es la del actor (`actors`), que la plataforma
        firma y nadie puede falsificar desde la interfaz."""
        state = merge_state(author="ardelperal", merged_by="ardelperal",
                            message="Merge PR #7 (governed merge)\n\n"
                                    "ci: el titulo lo escribio el humano en la web")
        proc = run_gate(self.root, push_event(), governed_contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_without_the_key_the_verdict_does_not_change(self) -> None:
        """Sin la clave el detective se comporta como antes: el merge por la web de
        #373 no es violation (retrocompatibilidad)."""
        state = merge_state(author="ardelperal", merged_by="ardelperal", message=WEB_MESSAGE)
        proc = run_gate(self.root, push_event(), contract(), state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_a_disabled_key_does_not_change_the_verdict(self) -> None:
        """`enabled: false` es la clave presente pero apagada: el veredicto no cambia."""
        state = merge_state(author="ardelperal", merged_by="ardelperal", message=WEB_MESSAGE)
        contract_data = {**contract(), "governed_path": {"enabled": False, "actors": []}}
        proc = run_gate(self.root, push_event(), contract_data, state)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["created_issues"], [])

    def test_an_unusable_governed_path_fails_closed(self) -> None:
        """HR-3: un contrato inservible es duda, no un verde (exit 2)."""
        for bad in ({"enabled": "yes", "actors": ["github-actions[bot]"]},
                    {"enabled": True, "actors": []},
                    {"enabled": True, "actors": ["github-actions[bot]"],
                     "merge_marker": "["}):
            with self.subTest(governed_path=bad):
                proc = run_gate(self.root, push_event(),
                                {**contract(), "governed_path": bad}, base_state())
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertIn("governed_path", proc.stdout + proc.stderr)


def dispatch_event(before: object = BASE, after: object = SHA, **extra: object) -> dict:
    """#933: payload de `workflow_dispatch` con el rango en `inputs` (sin `commits` ni `before`)."""
    inputs = {"before": before, "after": after}
    return {"ref": "refs/heads/main", "inputs": inputs, "workflow": "push-compliance",
            "repository": {"full_name": REPO, "default_branch": "main"}, **extra}


class DispatchRangeTests(unittest.TestCase):
    """#933: el detective audita el rango `before...after` de un dispatch y falla cerrado sin él."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def stored(self) -> dict:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))

    def test_a_conforming_dispatch_audits_the_explicit_range_via_compare(self) -> None:
        proc = run_gate(self.root, dispatch_event(), contract(), base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn(f"GET repos/{REPO}/compare/{BASE}...{SHA}", self.stored()["calls"])
        self.assertIn("OK:", proc.stdout)

    def test_a_violating_commit_in_the_dispatched_range_opens_an_incident(self) -> None:
        state = base_state()
        state["pulls"] = {}  # ningún PR mergeado para el commit del rango
        proc = run_gate(self.root, dispatch_event(), contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(len(self.stored()["created_issues"]), 1)

    def test_a_dispatch_without_before_fails_closed_before_any_api_call(self) -> None:
        event = dispatch_event()
        del event["inputs"]["before"]
        proc = run_gate(self.root, event, contract(), base_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("before", proc.stderr)
        self.assertFalse(any(c.startswith("GET repos") and "compare" in c for c in self.stored()["calls"]))
        self.assertEqual(self.stored()["created_issues"], [])

    def test_a_dispatch_without_after_fails_closed(self) -> None:
        event = dispatch_event()
        del event["inputs"]["after"]
        proc = run_gate(self.root, event, contract(), base_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("after", proc.stderr)

    def test_a_dispatch_with_a_malformed_sha_fails_closed(self) -> None:
        proc = run_gate(self.root, dispatch_event(before="not-a-sha"), contract(), base_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.stored()["created_issues"], [])

    def test_a_dispatch_with_a_zero_before_fails_closed(self) -> None:
        proc = run_gate(self.root, dispatch_event(before="0" * 40), contract(), base_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_a_dispatch_whose_inputs_are_not_an_object_fails_closed(self) -> None:
        event = dispatch_event()
        event["inputs"] = "before...after"
        proc = run_gate(self.root, event, contract(), base_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main()
