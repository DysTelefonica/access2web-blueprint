#!/usr/bin/env python3
# ci-pattern asset tests — post-push compliance detective (DysTelefonica/team-skills#303)
"""Post-push detective suite (black-box): a violating push opens an issue
with SHA, author and check evidence (exit 1); a broken environment fails
closed (exit 2). ``gh`` is simulated with the real API shapes; stdlib only."""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent / "check_push_compliance.py"
REPO = "DysTelefonica/team-skills"
SHA = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
BASE = "0a1b2c3d4e5f60718293a4b5c6d7e8f90a1b2c3d"
CHECK = "unit suite (≤3min)"

# Simulacro de `gh` con las rutas reales de la API; path del estado incrustado (HR-28).
GH_SHIM = '''#!/usr/bin/env python3
import json, sys
STATE_PATH = {state_path!r}
state = json.load(open(STATE_PATH, encoding="utf-8"))
args = sys.argv[1:]
state["calls"].append(" ".join(args))
if state.get("api_error"):
    sys.stderr.write("gh: HTTP 500 (api_error fixture)\\n"); sys.exit(1)
if "-X" in args and "POST" in args:
    if state.get("issue_post_error"):
        sys.stderr.write("gh: HTTP 503 (issue_post_error fixture)\\n"); sys.exit(1)
    endpoint = next((a for a in args if a.startswith("repos/") and a.endswith("/issues")), None)
    if endpoint is None:
        sys.stderr.write("gh: unsupported POST endpoint\\n"); sys.exit(1)
    fields = {{k: v for i, a in enumerate(args) if a in ("-f", "-F") and i + 1 < len(args)
               for k, _, v in [args[i + 1].partition("=")]}}
    state["created_issues"].append(fields)
    json.dump(state, open(STATE_PATH, "w", encoding="utf-8"))
    print(json.dumps({{"number": 42, "html_url": "https://github.com/" + endpoint.split("/issues")[0].replace("repos/", "") + "/issues/42", "title": fields.get("title", "")}}))
    sys.exit(0)
if "-X" in args and "PUT" in args:
    endpoint = next((a for a in args if a.startswith("repos/") and a.endswith("/merge")), None)
    if endpoint is None:
        sys.stderr.write("gh: unsupported PUT endpoint\\n"); sys.exit(1)
    if state.get("merge_rejected"):
        sys.stderr.write("gh: HTTP 405 (merge_rejected fixture)\\n"); sys.exit(1)
    if state.get("head_changed") and any(a.startswith("sha=") for a in args):
        sys.stderr.write("gh: HTTP 409 (Head branch was modified)\\n"); sys.exit(1)
    state["merges"].append(endpoint)
    json.dump(state, open(STATE_PATH, "w", encoding="utf-8"))
    print(json.dumps({{"sha": "2c7d9b1a4e", "merged": True, "message": "Pull Request successfully merged"}}))
    sys.exit(0)
path = next((a for a in args if a.startswith("repos/")), None)
if path is None:
    sys.stderr.write("gh: no endpoint\\n"); sys.exit(1)
if path.endswith("/pulls") and "/commits/" in path:
    sha = path.split("/commits/")[1].split("/")[0]
    if state.get("invalid_body"):
        print("{{not valid json"); sys.exit(0)
    pages = state.get("pulls_pages", {{}}).get(sha)
    if pages is not None:
        print(" ".join(json.dumps(page) for page in pages))
    else:
        print(json.dumps(state["pulls"].get(sha, [])))
elif "/check-runs" in path:
    sha = path.split("/commits/")[1].split("/")[0]
    runs = state["check_runs"].get(sha, [])
    print(json.dumps({{"total_count": len(runs), "check_runs": runs}}))
elif "/compare/" in path:
    print(json.dumps(state.get("compare", {{"total_commits": 0, "commits": []}})))
elif "/pulls/" in path:
    detail = state["pr_details"].get(path.rsplit("/pulls/", 1)[1])
    if detail is None:
        sys.stderr.write("gh: 404 Not Found\\n"); sys.exit(1)
    print(json.dumps(detail))
else:
    sys.stderr.write("gh: unsupported endpoint " + path + "\\n"); sys.exit(1)
json.dump(state, open(STATE_PATH, "w", encoding="utf-8"))
'''


def write_fixture(root: Path, event: dict, contract: dict, state: dict) -> None:
    (root / "event.json").write_text(json.dumps(event), encoding="utf-8")
    (root / "contract.json").write_text(json.dumps(contract), encoding="utf-8")
    (root / "state.json").write_text(json.dumps(state), encoding="utf-8")
    shim = root / "bin" / "gh"
    shim.parent.mkdir(exist_ok=True)
    shim.write_text(GH_SHIM.format(state_path=str(root / "state.json")), encoding="utf-8")
    shim.chmod(0o755)


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


def base_state() -> dict:
    return {"pulls": {SHA: [{"number": 7, "merged": True, "merged_at": "2026-10-06T10:00:00Z",
                             "user": {"login": "ardelperal"}, "base": {"ref": "main"}}]},
            "pr_details": {"7": {"number": 7, "merged": True, "user": {"login": "ardelperal"},
                                 "head": {"sha": SHA}, "merge_commit_sha": SHA}},
            "check_runs": {SHA: [{"id": 1, "name": CHECK, "conclusion": "success", "status": "completed"}]},
            "created_issues": [], "calls": [], "merges": [],
            "compare": {"total_commits": 1,
                        "commits": [{"sha": SHA,
                                     "commit": {"author": {"name": "Dev A"}}}]}}


def run_gate(root: Path, event: dict, contract_data: dict, state: dict) -> subprocess.CompletedProcess:
    write_fixture(root, event, contract_data, state)
    env = {**os.environ, "PATH": f"{root / 'bin'}{os.pathsep}{os.environ.get('PATH', '')}"}
    return subprocess.run([sys.executable, str(ASSET), "--event", str(root / "event.json"),
                           "--contract", str(root / "contract.json")],
                          capture_output=True, text=True, env=env, timeout=60)


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

    def test_commit_without_merged_pr_opens_incident(self) -> None:
        state = base_state()
        state["pulls"][SHA] = [{"number": 7, "merged": False, "merged_at": None,
                                "user": {"login": "ardelperal"}}]
        proc = run_gate(self.root, push_event(), contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        issue = self.created_issue()
        self.assertIn(SHA, issue["body"])
        self.assertIn("Dev A", issue["body"])
        self.assertIn("PR", issue["title"] + issue["body"])
        self.assertIn("INCIDENT ISSUE", proc.stdout + proc.stderr)

    def test_merged_pr_with_red_check_opens_incident(self) -> None:
        state = base_state()
        state["check_runs"][SHA] = [{"id": 1, "name": CHECK, "conclusion": "failure", "status": "completed"}]
        proc = run_gate(self.root, push_event(), contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        issue = self.created_issue()
        self.assertIn(CHECK, issue["body"])
        self.assertIn("failure", issue["body"])
        self.assertIn(SHA, issue["body"])

    def test_merged_pr_missing_required_check_is_a_violation(self) -> None:
        state = base_state()
        state["check_runs"][SHA] = []
        proc = run_gate(self.root, push_event(), contract(), state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        issue = self.created_issue()
        self.assertIn(CHECK, issue["body"])
        self.assertIn("missing", issue["body"])

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


if __name__ == "__main__":
    unittest.main()
