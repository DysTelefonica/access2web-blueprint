#!/usr/bin/env python3
"""Tests de nightly_pipeline.py (#1627, slice 3; #1691).

La API se lee por HTTP contra el servidor stub de runner-controls (GITHUB_API_URL
apunta a el), no con el CLI `gh`: el runner del consumer no lo trae. El PATH va
vacio en todas las pruebas, asi que un `gh` presente en la maquina no puede
salvarla.
"""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSETS = Path(__file__).resolve().parents[1]
ASSET = ASSETS / "nightly_pipeline.py"
sys.path.insert(0, str(ASSETS))
sys.path.insert(0, str(ASSETS.parent / "runner-controls" / "tests"))
import nightly_pipeline as np  # noqa: E402
from _github_stub import github_stub, stub_env  # noqa: E402

REPO = "DysTelefonica/team-skills"
RUN = 4242
SHA = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
NEW = "9a8b7c6d5e4f3a2b1c0d9e8f7a6b5c4d3e2f1a0b"
GREEN = {job: "success" for job in np.REQUIRED_JOBS}
PROFILE_OK = {"schema_version": 1, "enabled": True}
RUN_PATH = f"repos/{REPO}/actions/runs/{RUN}"
JOBS_PATH = f"{RUN_PATH}/jobs"
RUNS_PATH = f"repos/{REPO}/actions/runs"
#: La revision desplegada en produccion y un sha verde mas antiguo que la punta.
DEPLOYED = "5c4b3a2918f7e6d5c4b3a291807f6e5d4c3b2a19"
CANDIDATE = "2b1a0918273645546372819aabbccddeeff0011"


def jobs_payload(conclusions: dict) -> dict:
    return {"jobs": [{"name": name, "status": "completed", "conclusion": value}
                     for name, value in conclusions.items()]}


def base_state(run_sha: str = SHA, main_sha: str = SHA, jobs: dict | None = None) -> dict:
    return {"calls": [],
            "api": {RUN_PATH: {"id": RUN, "head_sha": run_sha},
                    JOBS_PATH: jobs_payload(jobs if jobs is not None else GREEN)},
            "base_tips": {"main": main_sha}}


#: La noche de un consumer con sus propios nombres (#1699): `e2e` en lugar de `e2e-full`
#: y `trivy` dentro de `build`.
CONSUMER_JOBS = {"build": "success", "e2e": "success", "integration": "success",
                 "upgrade-path": "success", "security-deep": "success"}
JOB_MAP = ("--job", "e2e-full=e2e", "--job", "trivy=build")


DECIDE_CASES = [
    ("todo en verde sobre la cabeza", SHA, SHA, GREEN, 0, True, [SHA[:12]]),
    ("main se ha movido", SHA, NEW, GREEN, 0, False, [f"main is now {NEW[:12]}", "re-dispatch"]),
    ("un job en fallo", SHA, SHA, {**GREEN, "e2e-full": "failure"}, 0, False, ["e2e-full=failure"]),
    ("un job sin terminar", SHA, SHA, {**GREEN, "trivy": None}, 0, False, ["trivy=None"]),
    ("falta un job requerido", SHA, SHA, {k: v for k, v in GREEN.items() if k != "upgrade-path"}, 0,
     False, ["upgrade-path missing"]),
    ("migracion pendiente", SHA, SHA, GREEN, 2, False, ["2 pending migration", "never applies one alone"]),
]

PROFILE_CASES = [
    ("activo", {"schema_version": 1, "enabled": True}, True),
    ("desactivado", {"schema_version": 1, "enabled": False}, False),
    ("JSON no valido", "{not json", None),
    ("esquema 2", {"schema_version": 2, "enabled": True}, None),
    ("enabled no booleano", {"schema_version": 1, "enabled": "yes"}, None),
    ("no es un objeto", [], None),
]


class DecideTest(unittest.TestCase):
    def test_decision_table(self) -> None:
        for name, run_sha, main_sha, jobs, pending, deploy, parts in DECIDE_CASES:
            with self.subTest(case=name):
                got, reason = np.decide(run_sha, main_sha, jobs, np.REQUIRED_JOBS, pending)
                self.assertEqual(got, deploy, reason)
                for part in parts:
                    self.assertIn(part, reason)


def green_run(run_id: int, sha: str, path: str = ".github/workflows/ci.yml") -> dict:
    """Una ejecucion en verde sobre ``sha`` (la evidencia de que ese sha esta en verde)."""
    return {"id": run_id, "head_sha": sha, "status": "completed", "conclusion": "success",
            "name": "ci", "path": path}


def compares(*specs) -> dict:
    """Las comparaciones de la API: cada rango ``base...head`` con su estado."""
    return {f"{base}...{head}": {"status": status} for base, head, status in specs}


def candidate_state(runs=None, compares=None, deployed=DEPLOYED, jobs=None) -> dict:
    """El estado del stub: la punta, la revision desplegada, la evidencia y las comparaciones."""
    state = base_state()
    state["api"][RUNS_PATH] = {"workflow_runs": [green_run(1, CANDIDATE)] if runs is None else runs}
    state["api"][f"repos/{REPO}/actions/runs/1/jobs"] = jobs_payload(GREEN if jobs is None else jobs)
    state["deployments"] = [{"id": 7, "sha": deployed, "environment": "production"}]
    state["deployment_statuses"] = {"7": [{"state": "success"}]}
    state["compares"] = {} if compares is None else compares
    return state


#: (nombre, evidencia, jobs de la evidencia, comparaciones, candidato esperado).
CANDIDATE_CASES = [
    ("el verde mas reciente que desciende de la punta",
     [green_run(1, CANDIDATE)], GREEN,
     compares((SHA, CANDIDATE, "behind"), (DEPLOYED, CANDIDATE, "ahead")), CANDIDATE),
    ("un sha fuera de la historia de la punta no es candidato",
     [green_run(1, CANDIDATE)], GREEN, compares((SHA, CANDIDATE, "ahead")), SHA),
    ("un candidato que no desciende de la revision desplegada se descarta",
     [green_run(1, CANDIDATE)], GREEN,
     compares((SHA, CANDIDATE, "behind"), (DEPLOYED, CANDIDATE, "behind")), SHA),
    ("sin evidencia verde el candidato es la punta",
     [], GREEN, {}, SHA),
    ("una ejecucion en rojo no es evidencia",
     [green_run(1, CANDIDATE)], {**GREEN, "e2e-full": "failure"},
     compares((SHA, CANDIDATE, "behind"), (DEPLOYED, CANDIDATE, "ahead")), SHA),
    ("las ejecuciones de la propia nocturna no son evidencia",
     [green_run(1, CANDIDATE, ".github/workflows/nightly-pipeline.yml")], GREEN,
     compares((SHA, CANDIDATE, "behind"), (DEPLOYED, CANDIDATE, "ahead")), SHA),
]

#: (nombre, comparaciones, jobs de la ejecucion, migraciones, codigo esperado, motivo esperado).
GATE_CASES = [
    ("el candidato verde que desciende de la revision desplegada despliega",
     compares((SHA, CANDIDATE, "behind"), (DEPLOYED, CANDIDATE, "ahead")), GREEN, "0", 0, "DEPLOY: "),
    ("un candidato que no desciende de la revision desplegada retiene",
     compares((SHA, CANDIDATE, "behind"), (DEPLOYED, CANDIDATE, "behind")), GREEN, "0", 2,
     "refusing to roll back"),
    ("un candidato fuera de la historia de la punta retiene",
     compares((SHA, CANDIDATE, "ahead")), GREEN, "0", 2, "not in the history of the main head"),
    ("un job en rojo sobre el candidato retiene",
     compares((SHA, CANDIDATE, "behind"), (DEPLOYED, CANDIDATE, "ahead")),
     {**GREEN, "trivy": "failure"}, "0", 2, "trivy=failure"),
    ("una migracion pendiente retiene",
     compares((SHA, CANDIDATE, "behind"), (DEPLOYED, CANDIDATE, "ahead")), GREEN, "2", 2,
     "2 pending migration"),
]


class ProfileTest(unittest.TestCase):
    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.path = Path(tmp.name) / "prerelease.json"

    def test_profile_table(self) -> None:
        for name, body, enabled in PROFILE_CASES:
            with self.subTest(case=name):
                self.path.write_text(body if isinstance(body, str) else json.dumps(body), encoding="utf-8")
                if enabled is None:
                    with self.assertRaises(np.ProfileError):
                        np.read_profile(self.path)
                else:
                    self.assertIs(np.read_profile(self.path), enabled)

    def test_missing_profile_fails_closed(self) -> None:
        with self.assertRaises(np.ProfileError):
            np.read_profile(self.path)


class PlanCliTest(unittest.TestCase):
    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.profile = self.root / "prerelease.json"
        self.profile.write_text(json.dumps(PROFILE_OK), encoding="utf-8")

    def plan(self, state: dict, pending: str = "0", extra: tuple = ()):
        """Ejecuta `plan` contra el servidor stub; devuelve (proceso, llamadas)."""
        with tempfile.TemporaryDirectory() as tmp:
            state_file = Path(tmp) / "state.json"
            with github_stub(state_file, state) as base:
                env = stub_env(base)
                # Sin `gh` en el PATH: el asset no puede depender de el (#1691).
                env["PATH"] = ""
                args = ["plan", "--repo", REPO, "--run", str(RUN), "--profile", str(self.profile),
                        "--pending-migrations", pending, *extra]
                proc = subprocess.run([sys.executable, str(ASSET), *args], env=env,  # noqa: S603
                                      capture_output=True, text=True, encoding="utf-8", timeout=60)
        return proc, list(state["calls"])

    def test_runs_without_gh_on_path(self) -> None:
        """El PATH va vacio: el resultado viene de la API, no de un `gh` (#1691)."""
        proc, calls = self.plan(base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertTrue(proc.stdout.splitlines()[-1].startswith("DEPLOY: "), proc.stdout)
        self.assertTrue(calls, "plan must read the run, the jobs and main")

    def test_prefixed_job_of_another_caller_does_not_count_as_the_required_job(self) -> None:
        for bare, prefixed in (("security-deep", "pr-checks / security-deep"), ("e2e-full", "build / e2e-full"),
                               ("build", "x / y / build")):
            with self.subTest(job=bare):
                jobs = [(n, "completed", "success") for n in np.REQUIRED_JOBS if n != bare]
                jobs.append((prefixed, "completed", "success"))
                state = base_state()
                state["api"][JOBS_PATH] = {"jobs": [
                    {"name": n, "status": s, "conclusion": c} for n, s, c in jobs]}
                proc, _ = self.plan(state)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertIn(f"{bare} missing", proc.stdout)

    def test_with_caller_other_prefixes_do_not_count(self) -> None:
        for name in ("other / trivy", "nightly / other / trivy"):
            with self.subTest(name=name):
                jobs = [(n, "completed", "success") for n in np.REQUIRED_JOBS if n != "trivy"]
                state = base_state()
                state["api"][JOBS_PATH] = {"jobs": [
                    {"name": n, "status": s, "conclusion": c} for n, s, c in [*jobs, (name, "completed", "success")]]}
                proc, _ = self.plan(state, extra=("--caller", "nightly"))
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertIn("trivy missing", proc.stdout)

    def test_the_single_caller_given_with_caller_is_accepted(self) -> None:
        state = base_state()
        state["api"][JOBS_PATH] = {"jobs": [
            {"name": f"nightly / {n}", "status": "completed", "conclusion": "success"} for n in np.REQUIRED_JOBS]}
        proc, _ = self.plan(state, extra=("--caller", "nightly"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertTrue(proc.stdout.splitlines()[-1].startswith("DEPLOY: "), proc.stdout)

    def pages(self, *pages) -> dict:
        state = base_state()
        del state["api"][JOBS_PATH]
        state["api_pages"] = {JOBS_PATH: list(pages)}
        return state

    def test_jobs_read_follows_the_link_header(self) -> None:
        first = {"jobs": [{"name": n, "status": "completed", "conclusion": "success"} for n in np.REQUIRED_JOBS[:3]]}
        second = {"jobs": [{"name": n, "status": "completed", "conclusion": "success"} for n in np.REQUIRED_JOBS[3:]]}
        proc, calls = self.plan(self.pages(first, second))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertTrue(proc.stdout.splitlines()[-1].startswith("DEPLOY: "), proc.stdout)
        self.assertEqual([call for call in calls if call.endswith(JOBS_PATH)], [f"GET {JOBS_PATH}"] * 2)

    def test_failure_on_a_later_page_holds(self) -> None:
        first = {"jobs": [{"name": n, "status": "completed", "conclusion": "success"} for n in np.REQUIRED_JOBS[:3]]}
        second = {"jobs": [{"name": n, "status": "completed", "conclusion": "success"} for n in np.REQUIRED_JOBS[3:]]}
        second["jobs"][0]["conclusion"] = "failure"
        proc, _ = self.plan(self.pages(first, second))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("upgrade-path=failure", proc.stdout)

    def test_bad_page_fails_closed(self) -> None:
        good = {"jobs": [{"name": n, "status": "completed", "conclusion": "success"} for n in np.REQUIRED_JOBS]}
        for bad in ({"total_count": 1}, "not-json-page"):
            with self.subTest(page=bad):
                proc, _ = self.plan(self.pages(good, bad))
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertTrue(proc.stdout.splitlines()[-1].startswith("HOLD: "), proc.stdout)

    def test_deploys_when_green_on_the_head(self) -> None:
        proc, _ = self.plan(base_state())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertTrue(proc.stdout.splitlines()[-1].startswith("DEPLOY: "), proc.stdout)

    def test_holds_when_main_moved(self) -> None:
        proc, _ = self.plan(base_state(main_sha=NEW))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertTrue(proc.stdout.splitlines()[-1].startswith("HOLD: "), proc.stdout)

    def test_is_read_only(self) -> None:
        _, calls = self.plan(base_state(), pending="1")
        self.assertTrue(calls, "plan must read the run, the jobs and main")
        self.assertTrue(all(call.startswith("GET ") for call in calls), calls)

    def test_disabled_profile_holds(self) -> None:
        self.profile.write_text(json.dumps({"schema_version": 1, "enabled": False}), encoding="utf-8")
        proc, _ = self.plan(base_state())
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("profile", proc.stdout.splitlines()[-1])

    def entries(self, before: tuple = (), after: tuple = (), replace: str = "", status: str = "completed",
                prefix: str = "", conclusion: str = "success") -> dict:
        green = [(job, "completed", "success") for job in np.REQUIRED_JOBS if job != replace]
        if replace:
            green.append((replace, status, conclusion))
        state = base_state()
        state["api"][JOBS_PATH] = {"jobs": [
            {"name": prefix + n, "status": s, "conclusion": c} for n, s, c in [*before, *green, *after]]}
        return state

    def test_workflow_call_prefixed_names_with_the_caller_deploy_or_hold(self) -> None:
        proc, _ = self.plan(self.entries(prefix="nightly / "), extra=("--caller", "nightly"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        proc, _ = self.plan(self.entries(prefix="nightly / ", replace="trivy", conclusion="failure"),
                            extra=("--caller", "nightly"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("trivy=failure", proc.stdout)

    def test_empty_pending_value_is_refused_without_api_calls(self) -> None:
        proc, calls = self.plan(self.entries(), pending="")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(calls, [])

    def test_repeated_job_with_a_failure_in_either_order_holds(self) -> None:
        for kwargs in ({"before": (("e2e-full", "completed", "failure"),)},
                       {"after": (("e2e-full", "completed", "failure"),)}):
            with self.subTest(order=list(kwargs)):
                proc, _ = self.plan(self.entries(**kwargs))
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertIn("e2e-full=failure", proc.stdout)

    def test_job_not_completed_is_not_green(self) -> None:
        proc, _ = self.plan(self.entries(replace="trivy", status="in_progress"))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("trivy=None", proc.stdout)

    def test_plan_twice_gives_identical_output_code_and_calls(self) -> None:
        first = self.plan(base_state())
        second = self.plan(base_state())
        self.assertEqual(first[0].stdout, second[0].stdout)
        self.assertEqual(first[0].returncode, second[0].returncode)
        self.assertEqual(first[1], second[1])

    def test_unreadable_api_fails_closed(self) -> None:
        state = base_state()
        del state["api"][JOBS_PATH]
        proc, _ = self.plan(state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertTrue(proc.stdout.splitlines()[-1].startswith("HOLD: "), proc.stdout)
        self.assertIn("redispatch_nightly.py --repo", proc.stdout)

    def test_negative_pending_is_refused_without_calls(self) -> None:
        proc, calls = self.plan(base_state(), pending="-1")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(calls, [])

    # `--job <requerido>=<nombre del consumer>` (#1699).
    def test_mapping_all_green_deploys(self) -> None:
        proc, calls = self.plan(base_state(jobs=CONSUMER_JOBS), extra=JOB_MAP)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(proc.stdout, f"DEPLOY: all 6 jobs green on {SHA[:12]}\n")
        self.assertEqual(proc.stderr, "")
        self.assertTrue(calls)

    def test_mapped_job_red_holds_naming_the_required_job(self) -> None:
        proc, _ = self.plan(base_state(jobs={**CONSUMER_JOBS, "e2e": "failure"}), extra=JOB_MAP)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(proc.stdout, "HOLD: not all green: e2e-full=failure\n")

    def test_two_required_jobs_mapped_to_one_red_job_both_hold(self) -> None:
        proc, _ = self.plan(base_state(jobs={**CONSUMER_JOBS, "build": "failure"}), extra=JOB_MAP)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(proc.stdout, "HOLD: not all green: build=failure, trivy=failure\n")

    def test_mapping_goes_through_the_caller_prefix(self) -> None:
        state = base_state(jobs={f"nightly / {n}": c for n, c in CONSUMER_JOBS.items()})
        proc, _ = self.plan(state, extra=(*JOB_MAP, "--caller", "nightly"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_without_mapping_consumer_names_do_not_count(self) -> None:
        proc, _ = self.plan(base_state(jobs=CONSUMER_JOBS))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(proc.stdout,
                         "HOLD: required job not in the run: e2e-full missing; trivy missing\n")

    def test_invalid_mapping_is_refused_without_api_calls(self) -> None:
        cases = [
            ("requerido desconocido", ("--job", "e2e=e2e"), "HOLD: --job e2e=e2e: unknown required job 'e2e'"),
            ("requerido repetido", ("--job", "trivy=build", "--job", "trivy=scan"),
             "HOLD: --job trivy=scan: required job 'trivy' is already mapped"),
            ("sin '='", ("--job", "trivy"), "HOLD: --job trivy: expected <required>=<consumer job>"),
            ("nombre vacio", ("--job", "trivy="), "HOLD: --job trivy=: expected <required>=<consumer job>"),
            ("requerido vacio", ("--job", "=build"), "HOLD: --job =build: expected <required>=<consumer job>"),
        ]
        for name, extra, message in cases:
            with self.subTest(case=name):
                proc, calls = self.plan(base_state(), extra=extra)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertEqual(proc.stdout, message + "\n")
                self.assertEqual(proc.stderr, "")
                self.assertEqual(calls, [])


class CandidateCliTest(unittest.TestCase):
    """La eleccion del sha de deploy: el verde mas reciente que descienda de la punta (#1724)."""

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.profile = Path(tmp.name) / "prerelease.json"
        self.profile.write_text(json.dumps(PROFILE_OK), encoding="utf-8")

    def candidate(self, state: dict, extra: tuple = ()):
        """Ejecuta `candidate` contra el servidor stub; devuelve (proceso, llamadas)."""
        with tempfile.TemporaryDirectory() as tmp:
            state_file = Path(tmp) / "state.json"
            with github_stub(state_file, state) as base:
                env = stub_env(base)
                env["PATH"] = ""
                args = ["candidate", "--repo", REPO, "--profile", str(self.profile), *extra]
                proc = subprocess.run([sys.executable, str(ASSET), *args], env=env,  # noqa: S603
                                      capture_output=True, text=True, encoding="utf-8", timeout=60)
                return proc, list(state["calls"])

    def test_the_candidate_table(self) -> None:
        for name, runs, jobs, cmp, expected in CANDIDATE_CASES:
            with self.subTest(case=name):
                proc, _ = self.candidate(candidate_state(runs=runs, jobs=jobs, compares=cmp))
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
                self.assertIn(f"CANDIDATE: {expected}", proc.stdout)

    def test_it_names_the_deployed_revision(self) -> None:
        state = candidate_state(compares=compares((SHA, CANDIDATE, "behind"),
                                                  (DEPLOYED, CANDIDATE, "ahead")))
        proc, _ = self.candidate(state)
        self.assertIn(f"DEPLOYED: {DEPLOYED}", proc.stdout)

    def test_an_unreadable_state_holds_with_a_named_reason(self) -> None:
        state = candidate_state()
        state["deployments_error"] = True
        proc, _ = self.candidate(state)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertTrue(proc.stdout.startswith("HOLD: "), proc.stdout)


class PlanCandidateTest(unittest.TestCase):
    """El gate con candidato: en la historia de la punta, descendiente de la desplegada y verde."""

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.profile = Path(tmp.name) / "prerelease.json"
        self.profile.write_text(json.dumps(PROFILE_OK), encoding="utf-8")

    def plan(self, state: dict, pending: str = "0"):
        """Ejecuta `plan --candidate` contra el servidor stub; devuelve (proceso, llamadas)."""
        with tempfile.TemporaryDirectory() as tmp:
            state_file = Path(tmp) / "state.json"
            with github_stub(state_file, state) as base:
                env = stub_env(base)
                env["PATH"] = ""
                args = ["plan", "--repo", REPO, "--run", str(RUN), "--profile", str(self.profile),
                        "--pending-migrations", pending, "--candidate", CANDIDATE]
                proc = subprocess.run([sys.executable, str(ASSET), *args], env=env,  # noqa: S603
                                      capture_output=True, text=True, encoding="utf-8", timeout=60)
                return proc, list(state["calls"])

    def test_the_gate_table(self) -> None:
        for name, cmp, jobs, pending, code, reason in GATE_CASES:
            with self.subTest(case=name):
                state = candidate_state(compares=cmp)
                state["api"][JOBS_PATH] = jobs_payload(jobs)
                proc, _ = self.plan(state, pending=pending)
                self.assertEqual(proc.returncode, code, proc.stdout + proc.stderr)
                self.assertIn(reason, proc.stdout.splitlines()[-1])


if __name__ == "__main__":
    unittest.main()
