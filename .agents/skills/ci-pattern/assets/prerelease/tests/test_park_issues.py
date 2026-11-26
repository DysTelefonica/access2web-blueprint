#!/usr/bin/env python3
# ci-pattern asset tests -- aparcar y desaparcar issues durante una noche en rojo (DysTelefonica/team-skills#1627, slice 5)
"""``park_issues.py``: aparca las issues normales en una noche roja y las desaparca tras un deploy verde.

Los casos del estado nocturno son tablas. La API es el servidor stub
(``runner-controls/tests/_github_stub.py``) y el PATH no trae ningun ``gh``:
el runner del consumer no lo tiene (#1691).
"""
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path

ASSETS = Path(__file__).resolve().parents[1]
ASSET = ASSETS / "park_issues.py"
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "runner-controls" / "tests"))
from _github_stub import github_stub, stub_env  # noqa: E402

sys.path.insert(0, str(ASSETS))
import park_issues as pk  # noqa: E402

REPO = "DysTelefonica/team-skills"
NIGHT = 9001
FIX = pk.FIX_LABEL
PARK = pk.PARK_LABEL
NIGHTLY_PATH = ".github/workflows/nightly.yml"
TODAY = time.strftime("%Y-%m-%d", time.gmtime())


def issue(number: int, state: str = "OPEN", labels: tuple = ()) -> dict:
    return {"number": number, "state": state, "labels": [{"name": name} for name in labels]}


def night(conclusion, status: str = "completed", path: str = NIGHTLY_PATH, name: str = "nightly") -> dict:
    return {"id": NIGHT, "status": status, "conclusion": conclusion, "path": path, "name": name}


def deploy(number: int, environment: str = "production", day: str = TODAY) -> dict:
    return {"id": number, "environment": environment, "created_at": f"{day}T08:00:00Z"}


# (caso, estado nocturno, resultado esperado)
NIGHT_CASES = [
    ("exito", night("success"), "green"),
    ("fallo", night("failure"), "red"),
    ("cancelada", night("cancelled"), "red"),
    ("timed_out", night("timed_out"), "red"),
    ("todavia corriendo", night(None, "in_progress"), "unknown"),
    ("sin conclusion", night(None), "unknown"),
]


class NightStateTest(unittest.TestCase):
    def test_night_state_table(self) -> None:
        for name, run, expected in NIGHT_CASES:
            with self.subTest(case=name):
                self.assertEqual(pk.night_state(run), expected)

    def test_non_dict_run_is_unknown(self) -> None:
        self.assertEqual(pk.night_state([]), "unknown")


class ParkCliTest(unittest.TestCase):
    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.issues = [
            issue(1), issue(2, labels=(FIX,)), issue(3, labels=(PARK,)),
            issue(4, state="CLOSED"), issue(5),
        ]
        self.deployments = [deploy(7)]
        self.statuses = {7: ["success"]}
        self.issues_fail = False

    def state(self, night_run) -> dict:
        return {"calls": [], "label_edits": [], "label_attempts": [], "fail_edit": [],
                "night_run": night_run, "issue_list": self.issues,
                "deployments": self.deployments,
                "deployment_statuses": {str(k): [{"state": s} for s in v]
                                        for k, v in self.statuses.items()}}

    def run_cli(self, command: str, night_run, extra: tuple = (), fail_edit=(), timeout: int = 60,
                without_token: bool = False):
        """Corre el asset contra el stub, sin ningun ``gh`` en el PATH.

        ``night_run`` None hace que la ejecucion nocturna sea ilegible; ``self.issues_fail``
        hace que el listado de issues no se pueda leer; ``without_token`` quita el token
        de la API (falla cerrado).
        """
        state = self.state(night_run)
        state["fail_edit"] = list(fail_edit)
        if self.issues_fail:
            state["issue_list_error"] = True
        if self.deployments is None:
            state["deployments_error"] = True
            state["deployments"] = []
        empty_bin = self.root / "empty-bin"
        empty_bin.mkdir(exist_ok=True)
        with github_stub(self.root / "state.json", state) as base:
            env = stub_env(base)
            env["PATH"] = str(empty_bin)  # ningun gh: el runner del consumer no lo tiene
            if without_token:
                env.pop("GITHUB_TOKEN", None)
                env.pop("GH_TOKEN", None)
            args = [command, "--repo", REPO, "--night-run", str(NIGHT), *extra]
            proc = subprocess.run([sys.executable, str(ASSET), *args],
                                  capture_output=True, text=True, env=env, timeout=timeout,
                                  encoding="utf-8")
        self.attempts = state["label_attempts"]
        return proc, state["label_edits"]

    def last(self, proc) -> str:
        return proc.stdout.splitlines()[-1]

    def test_park_works_with_no_gh_on_the_path(self) -> None:
        """Sin `gh` en el PATH el asset sigue aparcando: habla con la API (#1691)."""
        proc, edits = self.run_cli("park", night("failure"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.last(proc), "PARKED: 2 issues")
        self.assertEqual(edits, [{"action": "add", "number": 1, "label": PARK},
                                 {"action": "add", "number": 5, "label": PARK}])

    def test_park_reads_the_lowercase_state_of_the_rest_api(self) -> None:
        """La API REST devuelve `state` en minusculas ("open"); el CLI `gh` lo daba en mayusculas."""
        self.issues = [{**i, "state": str(i["state"]).lower()} for i in self.issues]
        proc, edits = self.run_cli("park", night("failure"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.last(proc), "PARKED: 2 issues")
        self.assertEqual(edits, [{"action": "add", "number": 1, "label": PARK},
                                 {"action": "add", "number": 5, "label": PARK}])

    def test_park_adds_the_label_to_open_ordinary_issues_only(self) -> None:
        proc, edits = self.run_cli("park", night("failure"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(edits, [{"action": "add", "number": 1, "label": PARK},
                                 {"action": "add", "number": 5, "label": PARK}])
        self.assertEqual(self.last(proc), "PARKED: 2 issues")

    def test_park_dry_run_writes_nothing(self) -> None:
        proc, edits = self.run_cli("park", night("failure"), extra=("--dry-run",))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])
        self.assertTrue(self.last(proc).startswith("DRY-RUN: would park 2 issues: #1, #5"), proc.stdout)

    def test_park_refuses_unless_the_night_is_red(self) -> None:
        for label, run in (("green", night("success")), ("unknown", night(None, "in_progress")),
                           ("unreadable", None)):
            with self.subTest(night=label):
                proc, edits = self.run_cli("park", run)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertEqual(edits, [])
                self.assertTrue(self.last(proc).startswith("HOLD: "), proc.stdout)

    def test_park_holds_when_the_run_is_not_the_nightly_workflow(self) -> None:
        # Una ejecucion roja de `pr-size` no es la nocturna: no debe aparcar nada.
        proc, edits = self.run_cli("park", night("failure", path=".github/workflows/pr-size.yml", name="pr-size"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])
        self.assertTrue(self.last(proc).startswith("HOLD: "), proc.stdout)

    def test_workflow_match_is_exact(self) -> None:
        for path in (".github/workflows/nightly-old.yml", ".github/workflows/old-nightly.yml"):
            with self.subTest(path=path):
                proc, edits = self.run_cli("park", night("failure", path=path))
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertEqual(edits, [])

    def test_workflow_flag_names_the_workflow_to_accept(self) -> None:
        proc, edits = self.run_cli("park", night("failure", path=".github/workflows/pr-size.yml", name="pr-size"),
                                   extra=("--workflow", "pr-size"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(edits), 2)

    def test_park_partial_failure_exits_1_and_says_which(self) -> None:
        proc, edits = self.run_cli("park", night("failure"), fail_edit=(5,))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(edits, [{"action": "add", "number": 1, "label": PARK}])
        # La edicion rechazada se intenta igual y el resto no se detiene.
        self.assertEqual([a["number"] for a in self.attempts], [1, 5])
        self.assertTrue(self.last(proc).startswith("FAILED: "), proc.stdout)
        self.assertIn("#5", self.last(proc))

    def test_without_a_token_it_holds_and_edits_nothing(self) -> None:
        # Sin token la API no se puede leer: falla cerrado (salida 2), no edit nada.
        proc, edits = self.run_cli("park", night("failure"), without_token=True)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual((edits, self.attempts), ([], []))
        self.assertTrue(self.last(proc).startswith("HOLD: "), proc.stdout)

    def test_labels_are_overridable_by_flags(self) -> None:
        self.issues = [issue(1, labels=("custom:fix",)), issue(2)]
        proc, edits = self.run_cli("park", night("failure"),
                                   extra=("--fix-label", "custom:fix", "--park-label", "custom:parked"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(edits, [{"action": "add", "number": 2, "label": "custom:parked"}])

    def test_every_open_issue_is_parked_when_the_list_is_complete(self) -> None:
        self.issues = [issue(n) for n in range(1, 8)]
        proc, edits = self.run_cli("park", night("failure"), extra=("--list-limit", "10"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(edits), 7)
        self.assertEqual(self.last(proc), "PARKED: 7 issues")

    def test_list_exactly_at_the_limit_is_complete(self) -> None:
        self.issues = [issue(n) for n in range(1, 6)]
        proc, edits = self.run_cli("park", night("failure"), extra=("--list-limit", "5"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(edits), 5)

    def test_list_over_the_limit_holds_with_no_edits(self) -> None:
        self.issues = [issue(n) for n in range(1, 7)]
        proc, edits = self.run_cli("park", night("failure"), extra=("--list-limit", "5"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])
        self.assertTrue(self.last(proc).startswith("HOLD: "), proc.stdout)

    def test_list_limit_must_be_a_positive_integer(self) -> None:
        self.issues = []
        for value in ("0", "-1", "abc"):
            proc, edits = self.run_cli("park", night("failure"), extra=("--list-limit", value))
            self.assertEqual((proc.returncode, edits), (2, []), value)

    def test_unreadable_issue_list_holds_with_no_edits(self) -> None:
        # Una lista ilegible no es una lista vacia: sale 2 y no edita nada.
        self.issues_fail = True
        proc, edits = self.run_cli("park", night("failure"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])
        self.assertTrue(self.last(proc).startswith("HOLD: "), proc.stdout)

    def test_unpark_removes_the_label_from_every_state_after_a_green_night(self) -> None:
        self.issues = [issue(1, labels=(PARK,)), issue(2, state="CLOSED", labels=(PARK,)), issue(3)]
        proc, edits = self.run_cli("unpark", night("success"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(edits, [{"action": "remove", "number": 1, "label": PARK},
                                 {"action": "remove", "number": 2, "label": PARK}])
        self.assertEqual(self.last(proc), "UNPARKED: 2 issues")

    def test_unpark_refuses_unless_the_night_succeeded(self) -> None:
        self.issues = [issue(1, labels=(PARK,))]
        for label, run in (("red", night("failure")), ("unknown", night(None, "in_progress")),
                           ("unreadable", None)):
            with self.subTest(night=label):
                proc, edits = self.run_cli("unpark", run)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertEqual(edits, [])
                self.assertTrue(self.last(proc).startswith("HOLD: "), proc.stdout)

    def test_unpark_holds_without_a_production_deployment_today(self) -> None:
        # Un verde nocturno no basta: el desaparcado exige un deploy exitoso a produccion hoy.
        self.issues = [issue(1, labels=(PARK,))]
        self.deployments = [deploy(7, day="2001-01-01")]
        proc, edits = self.run_cli("unpark", night("success"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])
        self.assertTrue(self.last(proc).startswith("HOLD: "), proc.stdout)

    def test_unpark_holds_when_the_latest_deployment_status_is_not_success(self) -> None:
        self.issues = [issue(1, labels=(PARK,))]
        self.statuses = {7: ["failure", "success"]}
        proc, edits = self.run_cli("unpark", night("success"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])

    def test_unpark_environment_flag_names_the_deployment_environment(self) -> None:
        self.issues = [issue(1, labels=(PARK,))]
        self.deployments = [deploy(7, environment="staging")]
        proc, edits = self.run_cli("unpark", night("success"), extra=("--environment", "staging"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(len(edits), 1)

    def test_unpark_unreadable_deployments_hold(self) -> None:
        self.issues = [issue(1, labels=(PARK,))]
        self.deployments = None  # el stub no devuelve despliegues
        proc, edits = self.run_cli("unpark", night("success"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])

    def test_unpark_holds_on_a_deploy_from_yesterday_at_2359(self) -> None:
        # Un deploy de ayer a las 23:59:59 UTC no cuenta como deploy de hoy.
        yesterday = time.strftime("%Y-%m-%d", time.gmtime(time.time() - 86400))
        self.issues = [issue(1, labels=(PARK,))]
        self.deployments = [{"id": 7, "environment": "production", "created_at": f"{yesterday}T23:59:59Z"}]
        proc, edits = self.run_cli("unpark", night("success"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])
        self.assertTrue(self.last(proc).startswith("HOLD: "), proc.stdout)

    def test_unpark_holds_when_the_listed_deployment_is_another_environment(self) -> None:
        # La API devuelve un deploy de staging para la consulta de produccion: no cuenta.
        self.issues = [issue(1, labels=(PARK,))]
        self.deployments = [deploy(7, environment="staging")]
        proc, edits = self.run_cli("unpark", night("success"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])

    def test_unpark_is_idempotent(self) -> None:
        self.issues = [issue(1), issue(2, state="CLOSED")]
        proc, edits = self.run_cli("unpark", night("success"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])
        self.assertEqual(self.last(proc), "UNPARKED: 0 issues")

    def test_unpark_dry_run_writes_nothing(self) -> None:
        self.issues = [issue(1, labels=(PARK,))]
        proc, edits = self.run_cli("unpark", night("success"), extra=("--dry-run",))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(edits, [])
        self.assertTrue(self.last(proc).startswith("DRY-RUN: would unpark 1 issues: #1"), proc.stdout)


class PaginationTest(unittest.TestCase):
    """La API devuelve las issues paginadas: el asset sigue la cabecera Link (#1691)."""

    def pages(self, *pages):
        calls = []

        def fake_read_page(path):
            calls.append(path)
            return pages[len(calls) - 1]

        original = pk._read_page
        pk._read_page = fake_read_page
        self.addCleanup(setattr, pk, "_read_page", original)
        return calls

    def test_it_follows_the_next_page_and_ignores_the_pull_requests(self) -> None:
        link = ('<https://api.github.com/repositories/1/issues?page=2&per_page=100&state=all>; rel="next"')
        pull_request = {"number": 2, "state": "OPEN", "labels": [], "pull_request": {"url": "u"}}
        calls = self.pages(([issue(1), pull_request], link), ([issue(3)], ""))
        issues = pk.list_issues("o/r", 10)
        self.assertEqual([i["number"] for i in issues], [1, 3])
        self.assertEqual(len(calls), 2)
        self.assertTrue(calls[0].startswith("repos/o/r/issues?"), calls[0])
        self.assertEqual(calls[1], "/repositories/1/issues?page=2&per_page=100&state=all")

    def test_it_holds_when_the_pages_go_over_the_limit(self) -> None:
        link = '<https://api.github.com/repositories/1/issues?page=2>; rel="next"'
        self.pages(([issue(n) for n in (1, 2, 3)], link), ([issue(n) for n in (4, 5, 6)], ""))
        with self.assertRaises(pk.ReadError):
            pk.list_issues("o/r", 5)


if __name__ == "__main__":
    unittest.main()
