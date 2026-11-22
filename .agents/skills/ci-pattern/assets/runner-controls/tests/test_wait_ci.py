#!/usr/bin/env python3
# ci-pattern asset tests -- esperar a la CI de una PR (DysTelefonica/team-skills#1502)
"""`wait_ci.py` nunca devuelve exito con checks ausentes o pendientes.

Contra el stub HTTP de GitHub (sin `gh`): sin checks, pendientes, en rojo, en
verde, el agregador nombrado que no llega y la suite que aun esta arrancando
cuando ya hay un run verde (team-skills #1500)."""
import json
import subprocess
import sys
import tempfile
import unittest
from datetime import UTC, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _github_stub import github_stub, stub_env  # noqa: E402

ASSET = Path(__file__).resolve().parent.parent / "wait_ci.py"
REPO = "DysTelefonica/team-skills"
SHA = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
FAST = ("--poll-seconds", "0.05")


def run(name: str, status: str = "completed", conclusion: str | None = "success") -> dict:
    return {"id": abs(hash(name)) % 10000, "name": name, "status": status, "conclusion": conclusion}


def suite(status: str) -> dict:
    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    return {"id": 1, "status": status, "app": {"slug": "github-actions"}, "created_at": now}


def state(**extra: object) -> dict:
    base = {"pr_details": {"7": {"number": 7, "state": "open", "head": {"sha": SHA}}},
            "check_runs": {SHA: []}, "calls": []}
    base.update(extra)
    return base


class WaitCiTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def wait(self, data: dict, *extra: str) -> subprocess.CompletedProcess:
        with github_stub(self.root / "state.json", data) as base:
            return subprocess.run([sys.executable, str(ASSET), "--repo", REPO, "--pr", "7", *FAST, *extra],
                                  capture_output=True, text=True, encoding="utf-8",
                                  env=stub_env(base), timeout=60)

    def test_no_checks_ever_is_not_green(self) -> None:
        proc = self.wait(state(), "--appear-seconds", "0.3")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(f"WAITING: no check runs yet for {SHA[:12]}", proc.stdout)
        self.assertIn(f"NOT STARTED: no check runs appeared for {SHA[:12]} within 0.3 s; a base change "
                      "does not start CI: push an empty commit", proc.stdout)
        self.assertNotIn("GREEN", proc.stdout)

    def test_waits_through_absent_and_pending_then_green(self) -> None:
        data = state(check_runs_script={SHA: [[], [run("unit", "in_progress", None)],
                                              [run("unit"), run("lint", conclusion="skipped")]]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("WAITING: 1 check run(s) pending: unit", proc.stdout)
        self.assertIn(f"GREEN: 2 check run(s) of {SHA[:12]} passed", proc.stdout)

    def test_a_red_run_is_exit_1_naming_it(self) -> None:
        data = state(check_runs={SHA: [run("unit"), run("e2e", conclusion="failure"),
                                       run("lint", conclusion="cancelled")]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("RED: e2e (failure), lint (cancelled)", proc.stdout)

    def test_named_aggregator_that_never_reports_times_out(self) -> None:
        data = state(check_runs={SHA: [run("unit")]})
        proc = self.wait(data, "--check", "required", "--timeout-seconds", "0.3")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("WAITING: required check(s) not reported yet: required", proc.stdout)
        self.assertIn("TIMEOUT: required check(s) not reported yet: required after 0.3 s", proc.stdout)
        self.assertNotIn("GREEN", proc.stdout)

    def test_a_green_run_with_a_suite_still_starting_is_not_green_yet(self) -> None:
        # team-skills #1500: pr-size ya en verde y la suite de `tests` en cola sin runs.
        data = state(check_runs_script={SHA: [[run("pr-size")], [run("pr-size"), run("unit")]]},
                     check_suites_script={SHA: [[suite("completed"), suite("queued")],
                                                [suite("completed"), suite("completed")]]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn(f"WAITING: 1 GitHub Actions check suite(s) still starting for {SHA[:12]}", proc.stdout)
        self.assertIn(f"GREEN: 2 check run(s) of {SHA[:12]} passed", proc.stdout)

    def test_pending_forever_times_out(self) -> None:
        data = state(check_runs={SHA: [run("unit", "queued", None)]})
        proc = self.wait(data, "--timeout-seconds", "0.3")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("TIMEOUT: 1 check run(s) pending: unit after 0.3 s", proc.stdout)

    def test_api_error_fails_closed(self) -> None:
        proc = self.wait(state(api_error=True))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_bad_arguments_are_a_usage_error(self) -> None:
        proc = subprocess.run([sys.executable, str(ASSET), "--repo", REPO, "--pr", "0"],
                              capture_output=True, text=True, timeout=30)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("--pr must be positive", proc.stderr)


class AdvisoryChecksTests(unittest.TestCase):
    """#1762: un check advisory no pone la CI en rojo ni la deja esperando."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def contract(self, patterns: object, required: list | None = None) -> None:
        path = self.root / ".github" / "host-contract.json"
        path.parent.mkdir(parents=True, exist_ok=True)
        declared = {"contract_version": 1, "advisory_checks": patterns}
        if required is not None:
            declared["required_checks"] = [{"name": name, "class": "documented-only"}
                                           for name in required]
        path.write_text(json.dumps(declared), encoding="utf-8")

    def raw_contract(self, raw: bytes) -> None:
        path = self.root / ".github" / "host-contract.json"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(raw)


    def wait(self, data: dict, *extra: str) -> subprocess.CompletedProcess:
        with github_stub(self.root / "state.json", data) as base:
            return subprocess.run([sys.executable, str(ASSET), "--repo", REPO, "--pr", "7", *FAST, *extra],
                                  capture_output=True, text=True, encoding="utf-8",
                                  env=stub_env(base), timeout=60, cwd=self.root)

    def test_un_espejo_en_rojo_no_ponela_ci_en_rojo(self) -> None:
        self.contract(["*-gh"])
        data = state(check_runs={SHA: [run("unit"), run("docs-gh", conclusion="failure")]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("GREEN", proc.stdout)

    def test_sin_declaracion_el_comportamiento_es_el_de_siempre(self) -> None:
        data = state(check_runs={SHA: [run("unit"), run("docs-gh", conclusion="failure")]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("RED: docs-gh (failure)", proc.stdout)

    def test_un_espejo_pendiente_no_deja_esperando(self) -> None:
        self.contract(["*-gh"])
        data = state(check_runs={SHA: [run("unit"), run("pr-gate-gh", "in_progress", None)]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("GREEN", proc.stdout)

    def test_un_check_nombrado_no_queda_tapado_por_un_patron_advisory(self) -> None:
        self.contract(["required*"])
        data = state(check_runs={SHA: [run("required", conclusion="failure")]})
        proc = self.wait(data, "--check", "required")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("RED: required (failure)", proc.stdout)

    def test_una_declaracion_invalida_falla_cerrado(self) -> None:
        self.contract(["*-gh", 3])
        data = state(check_runs={SHA: [run("unit")]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("advisory_checks", proc.stdout + proc.stderr)

    def test_un_contrato_que_no_es_utf8_falla_cerrado(self) -> None:
        self.raw_contract(b'{"advisory_checks": ["\xff"]}')
        data = state(check_runs={SHA: [run("unit")]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_solo_checks_advisory_no_es_verde(self) -> None:
        # Un unico run advisory no es un CI verde: se sigue esperando.
        self.contract(["*-gh"])
        data = state(check_runs={SHA: [run("docs-gh", conclusion="failure")]})
        proc = self.wait(data, "--timeout-seconds", "0.3")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertNotIn("GREEN", proc.stdout)

    def test_un_check_requerido_del_contrato_no_queda_tapado(self) -> None:
        # El contrato lo declara requerido: sigue bloqueando sin `--check`.
        self.contract(["required*"], required=["required"])
        data = state(check_runs={SHA: [run("required", conclusion="failure")]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("RED: required (failure)", proc.stdout)

    # #1795: un required_checks mal formado falla cerrado y nombra la entrada rota.

    def required_contract(self, entries: object) -> None:
        path = self.root / ".github" / "host-contract.json"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({"contract_version": 1, "advisory_checks": ["*-gh"],
                                    "required_checks": entries}), encoding="utf-8")

    def assert_fails_closed(self, entries: object, must_name: str) -> None:
        self.required_contract(entries)
        data = state(check_runs={SHA: [run("unit")]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn(must_name, proc.stdout + proc.stderr)
        self.assertNotIn("GREEN", proc.stdout)

    def test_entrada_que_no_es_objeto_falla_cerrado(self) -> None:
        self.assert_fails_closed(["unit"], "required_checks[0]")

    def test_entrada_sin_name_falla_cerrado(self) -> None:
        self.assert_fails_closed([{"class": "documented-only"}], "required_checks[0]")

    def test_entrada_con_name_vacio_falla_cerrado(self) -> None:
        self.assert_fails_closed([{"name": ""}], "required_checks[0]")

    def test_required_checks_no_lista_falla_cerrado(self) -> None:
        self.assert_fails_closed({"name": "unit"}, "required_checks")

    def test_contrato_requerido_valido_no_cambia_el_veredicto(self) -> None:
        self.required_contract([{"name": "unit", "class": "documented-only"}])
        data = state(check_runs={SHA: [run("unit"), run("docs-gh", conclusion="failure")]})
        proc = self.wait(data)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("GREEN", proc.stdout)


if __name__ == "__main__":
    unittest.main()
