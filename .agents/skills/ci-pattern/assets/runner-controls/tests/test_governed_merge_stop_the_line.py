#!/usr/bin/env python3
"""Extremo a extremo de la parada de línea en el merge gobernado (#1627).

Usa el stub de test_governed_merge (run_gate, base_state, pull) con rutas para el
listado de issues, las ejecuciones de la nocturna y el perfil en la rama por defecto.
"""

import base64
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _github_stub import stub_env  # noqa: E402
from _labelled_stub import labelled_github_stub  # noqa: E402
from test_governed_merge import base_state, contract, pull  # noqa: E402

ASSET = Path(__file__).resolve().parents[1] / "check_governed_merge.py"

#: Distingue «la clave no está» de «la clave está con valor null»: la segunda es una
#: forma inválida y falla cerrado; la primera deja las señales activas (#1784).
ABSENT = object()

PROFILE = ".github/prerelease.json"
REPO = "DysTelefonica/team-skills"
NIGHT_LABEL = "prerelease:night-red"
MAIN_RED_LABEL = "prerelease:main-red"
RUN_URL = "https://github.com/DysTelefonica/team-skills/actions/runs/777"
REDISPATCH = "gh workflow run tests.yml --repo DysTelefonica/team-skills --ref main"
TYPE_BUG = [{"name": "type:bug"}]


def by_label(state):
    """Las issues de `issue_list` son de la nocturna; main-red queda vacío salvo que el test lo fije."""
    if "issues_by_label" in state:
        return state
    return {**state, "issues_by_label": {NIGHT_LABEL: state.get("issue_list", []), MAIN_RED_LABEL: []}}


def run_labelled(root, contract_data, state, *extra):
    """Corre el gate contra el stub con el listado de issues filtrado por etiqueta (#1708)."""
    (root / "contract.json").write_text(json.dumps(contract_data), encoding="utf-8")
    with labelled_github_stub(root / "state.json", by_label(state)) as base:
        env = stub_env(base)
        env.pop("GITHUB_STEP_SUMMARY", None)
        return subprocess.run([sys.executable, str(ASSET), "--repo", REPO, "--pr", "7",
                               "--contract", str(root / "contract.json"), *extra],
                              capture_output=True, text=True, env=env, timeout=60)


def profile_entry(enabled=True):
    raw = json.dumps({"enabled": enabled}).encode("utf-8")
    return {"content": base64.b64encode(raw).decode("ascii"), "encoding": "base64",
            "path": PROFILE, "name": "prerelease.json", "type": "file"}


def night(conclusion="failure", event="schedule", status="completed", created_at="",
          display_title=""):
    return {"id": 777, "event": event, "status": status, "conclusion": conclusion,
            "html_url": RUN_URL, "created_at": created_at, "display_title": display_title}


def red_state(runs=None, issues=(11,), linked=(), profile=True, runs_by_branch=None, **pull_over):
    """Estado con una issue etiquetada abierta, nocturna fijada y enlaces del PR."""
    state = base_state()
    state["issue_list"] = [{"number": n, "state": "open"} for n in issues]
    if runs_by_branch is not None:
        state["runs_by_branch"] = runs_by_branch
    else:
        state["runs"] = [night()] if runs is None else runs
    state["closing_refs"] = [{"number": n, "state": "open",
                              "repository": {"nameWithOwner": repo}} for repo, n in linked]
    if profile:
        state.setdefault("workflow_files", {})[PROFILE] = profile_entry(True)
    state["pr_details"] = {"7": pull(labels=TYPE_BUG, **pull_over)}
    return state


class TestStopTheLineEndToEnd(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)

    def gate(self, state, *extra):
        return run_labelled(self.root, contract(), state, *extra)

    def merges(self):
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))["merges"]

    def test_night_and_main_red_together_both_block_end_to_end(self):
        """#1708: las dos señales se evalúan; un PR que solo enlaza la de la nocturna sigue bloqueado."""
        state = red_state(issues=(11,), linked=[(REPO, 11)])
        state["issues_by_label"] = {NIGHT_LABEL: [{"number": 11, "state": "open"}],
                                    MAIN_RED_LABEL: [{"number": 21, "state": "open"}]}
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])
        self.assertIn("#21", proc.stdout)
        self.assertIn(MAIN_RED_LABEL, proc.stdout)

    def test_red_night_blocks_ordinary_pr_end_to_end(self):
        proc = self.gate(red_state())
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])
        self.assertIn(REDISPATCH, proc.stdout)

    def test_linked_issue_read_failure_fails_closed(self):
        state = red_state()
        state["graphql_error"] = True
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])

    def test_chain_intermediate_on_feature_branch_is_judged_by_the_default_branch(self):
        cases = (
            ("main green, feature red", {"main": [night("success")], "feat/x": [night("failure")]}, 0, 1),
            ("main red, feature green", {"main": [night("failure")], "feat/x": [night("success")]}, 1, 0),
        )
        for label, by_branch, exit_code, merged in cases:
            with self.subTest(label=label):
                tmp = tempfile.TemporaryDirectory()
                self.addCleanup(tmp.cleanup)
                self.root = Path(tmp.name)
                state = red_state(runs_by_branch=by_branch, issues=(11,),
                                  base={"ref": "feat/x", "sha": "b" * 40})
                state["pr_details"]["7"]["labels"] = [{"name": "chain:partial"}]
                proc = self.gate(state)
                self.assertEqual(proc.returncode, exit_code, proc.stdout + proc.stderr)
                self.assertEqual(len(self.merges()), merged)

    def test_green_marked_redispatch_after_red_night_lifts_the_line(self):
        """#1702: --dispatch-marker pasa al gate; un dispatch en verde con ese título limpia el rojo."""
        runs = [night("success", event="workflow_dispatch", created_at="2026-11-14T02:00:00Z",
                      display_title="suite=full"),
                night("failure", created_at="2026-11-14T01:00:00Z")]
        proc = self.gate(red_state(runs=runs), "--dispatch-marker", "suite=full")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.merges()), 1)

    def test_green_light_redispatch_without_the_marker_does_not_lift_the_line(self):
        """#1702: un dispatch ligero en verde no limpia el rojo, con o sin marcador en el gate."""
        runs = [night("success", event="workflow_dispatch", created_at="2026-11-14T02:00:00Z",
                      display_title="suite=light"),
                night("failure", created_at="2026-11-14T01:00:00Z")]
        proc = self.gate(red_state(runs=runs), "--dispatch-marker", "suite=full")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])

    def test_thirty_one_push_runs_do_not_hide_a_red_night_without_profile(self):
        pushes = [night("success", event="push", created_at="2026-11-14T03:00:00Z")
                  for _ in range(31)]
        runs = pushes + [night("failure", created_at="2026-11-14T01:00:00Z")]
        proc = self.gate(red_state(runs=runs, profile=False))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])

    def test_red_schedule_newer_than_a_green_dispatch_still_blocks(self):
        runs = [night("failure", created_at="2026-11-14T03:00:00Z"),
                night("success", event="workflow_dispatch", created_at="2026-11-14T02:00:00Z")]
        proc = self.gate(red_state(runs=runs))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])

    def test_one_event_query_failure_holds_even_without_profile(self):
        state = red_state(runs=[night("failure")], profile=False)
        state["fail_event"] = "schedule"
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])

    def test_cross_repo_link_with_same_number_is_not_a_fix(self):
        state = red_state(issues=(12,), linked=[("DysTelefonica/otro", 12)])
        proc = self.gate(state)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])

    def test_rule_holds_without_profile(self):
        proc = self.gate(red_state(profile=False))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])

    def test_no_profile_no_nightly_passes(self):
        proc = self.gate(red_state(runs=[], profile=False))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.merges()), 1)


class TestMainRedThroughTheInterface(unittest.TestCase):
    """#1708 (interfaz de APAP): la llamada exacta de `.github/workflows/runner-merge.yml`.

    La señal de `prerelease:main-red` está ACTIVA POR DEFECTO: no hay flag nuevo que
    pasar. El estado de la nocturna no se fija aquí (sin ejecuciones ni perfil), así
    que el veredicto depende solo de la señal main-red.
    """

    APAP_ARGS = ["--repo", REPO, "--contract", ".github/host-contract.json", "--pr", "7",
                 "--max-lines", "400", "--nightly-workflow", "ci.yml",
                 "--dispatch-marker", "night-redispatch"]

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        (self.root / ".github").mkdir()
        (self.root / ".github" / "host-contract.json").write_text(
            json.dumps(contract()), encoding="utf-8")

    def gate(self, state):
        with labelled_github_stub(self.root / "state.json", state) as base:
            env = stub_env(base)
            env.pop("GITHUB_STEP_SUMMARY", None)
            return subprocess.run([sys.executable, str(ASSET), *self.APAP_ARGS], cwd=self.root,
                                  capture_output=True, text=True, env=env, timeout=60)

    def state(self, main_red_issues=(), linked=()):
        state = base_state()
        state["pr_details"] = {"7": pull(labels=TYPE_BUG)}
        state["issues_by_label"] = {
            NIGHT_LABEL: [],
            MAIN_RED_LABEL: [{"number": n, "state": "open"} for n in main_red_issues]}
        state["closing_refs"] = [{"number": n, "state": "open",
                                  "repository": {"nameWithOwner": REPO}} for n in linked]
        return state

    def merges(self):
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))["merges"]

    def test_a_non_linked_pr_is_blocked_by_an_open_main_red_issue(self):
        proc = self.gate(self.state(main_red_issues=[21]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])
        self.assertIn("prerelease:main-red", proc.stdout)
        self.assertIn("#21", proc.stdout)

    def test_the_linked_fix_passes(self):
        proc = self.gate(self.state(main_red_issues=[21], linked=[21]))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.merges()), 1)

    def test_no_open_main_red_issue_passes(self):
        proc = self.gate(self.state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.merges()), 1)


class TestStopTheLineSwitchThroughTheInterface(unittest.TestCase):
    """#1784: el consumer apaga cada señal de la parada de línea desde su contrato.

    La declaración vive en `.github/host-contract.json` (la misma ruta que el
    merge gobernado ya lee): la clave `stop_the_line` con un miembro booleano
    por señal (`night_red_enabled`, `main_red_enabled`). Sin la clave —o sin
    un miembro— la señal sigue activa (el comportamiento de hoy). Una señal
    apagada no se consulta y no bloquea; el merge lo dice por `stderr`. Una
    forma inválida falla cerrado (2).
    """

    APAP_ARGS = ["--repo", REPO, "--contract", ".github/host-contract.json", "--pr", "7",
                 "--max-lines", "400", "--nightly-workflow", "ci.yml",
                 "--dispatch-marker", "night-redispatch"]

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        (self.root / ".github").mkdir()

    def write_contract(self, declaration=ABSENT):
        data = contract()
        if declaration is not ABSENT:
            data["stop_the_line"] = declaration
        (self.root / ".github" / "host-contract.json").write_text(
            json.dumps(data), encoding="utf-8")

    def gate(self, state, declaration=ABSENT):
        self.write_contract(declaration)
        with labelled_github_stub(self.root / "state.json", state) as base:
            env = stub_env(base)
            env.pop("GITHUB_STEP_SUMMARY", None)
            return subprocess.run([sys.executable, str(ASSET), *self.APAP_ARGS], cwd=self.root,
                                  capture_output=True, text=True, env=env, timeout=60)

    def state(self, night_red_issues=(), main_red_issues=(), linked=()):
        """Issues abiertas por etiqueta: la nocturna en `NIGHT_LABEL`, la de main en `MAIN_RED_LABEL`.

        La nocturna está en rojo de verdad (una ejecución `schedule` completada en
        fallo), así que la señal nocturna solo deja de bloquear si el consumer la
        apaga de verdad.
        """
        state = base_state()
        state["pr_details"] = {"7": pull(labels=TYPE_BUG)}
        state["issues_by_label"] = {
            NIGHT_LABEL: [{"number": n, "state": "open"} for n in night_red_issues],
            MAIN_RED_LABEL: [{"number": n, "state": "open"} for n in main_red_issues]}
        state["runs"] = [night()]
        state["closing_refs"] = [{"number": n, "state": "open",
                                  "repository": {"nameWithOwner": REPO}} for n in linked]
        return state

    def merges(self):
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))["merges"]

    def label_queries(self):
        """Las etiquetas que el gate consultó de verdad (el stub las registra)."""
        return json.loads((self.root / "state.json").read_text(encoding="utf-8")).get(
            "label_queries", [])

    def api_calls(self):
        """Las llamadas a la API que el gate hizo (el stub las registra)."""
        return json.loads((self.root / "state.json").read_text(encoding="utf-8")).get("calls", [])

    def test_without_the_key_both_signals_still_block(self):
        """Sin la clave, el comportamiento es el de hoy: ambas señales bloquean."""
        proc = self.gate(self.state(night_red_issues=[11], main_red_issues=[21]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])
        self.assertIn(NIGHT_LABEL, proc.stdout)
        self.assertIn(MAIN_RED_LABEL, proc.stdout)

    def test_the_night_red_signal_switched_off_does_not_query_its_label(self):
        """Con `night_red_enabled` apagada, el gate no consulta `prerelease:night-red` ni bloquea."""
        proc = self.gate(self.state(night_red_issues=[11], main_red_issues=[21]),
                         {"night_red_enabled": False})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])
        self.assertNotIn(NIGHT_LABEL, self.label_queries())
        self.assertIn(MAIN_RED_LABEL, self.label_queries())
        self.assertIn(NIGHT_LABEL, proc.stderr)

    def test_the_main_red_signal_switched_off_does_not_query_its_label(self):
        """Con `main_red_enabled` apagada, el gate no consulta `prerelease:main-red` ni bloquea."""
        proc = self.gate(self.state(night_red_issues=[11], main_red_issues=[21]),
                         {"main_red_enabled": False})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])
        self.assertNotIn(MAIN_RED_LABEL, self.label_queries())
        self.assertIn(NIGHT_LABEL, self.label_queries())
        self.assertIn(MAIN_RED_LABEL, proc.stderr)

    def test_both_signals_switched_off_let_the_merge_through(self):
        """Con las dos señales apagadas, nada bloquea y el merge sale."""
        proc = self.gate(self.state(night_red_issues=[11], main_red_issues=[21]),
                         {"night_red_enabled": False, "main_red_enabled": False})
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(self.merges()), 1)
        self.assertEqual(self.label_queries(), [])

    def test_the_notice_appears_even_when_another_gate_blocks_first(self):
        """El aviso se dice siempre: si otra puerta bloquea, el operador sigue viendo el interruptor."""
        state = self.state(night_red_issues=[11])
        state["pr_details"] = {"7": pull(labels=TYPE_BUG, draft=True)}
        proc = self.gate(state, {"night_red_enabled": False})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.merges(), [])
        self.assertIn(NIGHT_LABEL, proc.stderr)

    def test_a_malformed_declaration_fails_closed(self):
        """Una forma inválida no se ignora en silencio: el merge falla cerrado (2)."""
        for declaration in ("night-red", {"night_red_enabled": "no"}, {"unknown_signal": False},
                            None, [], 0, {"night_red_enabled": None}):
            with self.subTest(declaration=declaration):
                proc = self.gate(self.state(night_red_issues=[11], main_red_issues=[21]),
                                  declaration)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertEqual(self.merges(), [])
                self.assertIn("stop_the_line", proc.stderr)

    def test_a_malformed_declaration_fails_before_any_api_call(self):
        """La forma se valida antes de la API: una declaración rota no gasta una llamada."""
        proc = self.gate(self.state(night_red_issues=[11]), {"night_red_enabled": "no"})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.label_queries(), [])
        self.assertEqual(self.api_calls(), [])


if __name__ == "__main__":
    unittest.main()
