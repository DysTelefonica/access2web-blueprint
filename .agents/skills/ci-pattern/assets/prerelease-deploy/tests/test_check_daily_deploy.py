#!/usr/bin/env python3
# ci-pattern asset tests — comprobación diaria de deploy (#1627 slice 7, #1691)
"""Suite de check_daily_deploy.py: debe haber al menos un deploy exitoso a
production por día UTC; si no lo hay, se abre una sola issue de máxima prioridad
con la etiqueta fija `prerelease:night-red`. El asset habla con la API por
`runner-controls/github_rest.py` (#1691): las pruebas apuntan GITHUB_API_URL al
servidor stub local y no tocan la red ni el CLI `gh` (el runner del consumer no
lo trae).

Ejecución: `python3 -m unittest discover -s tests -p 'test_check_daily_deploy*.py'`.
Solo stdlib."""

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parents[1]
SCRIPT = ASSET / "check_daily_deploy.py"
sys.path.insert(0, str(ASSET.parent / "runner-controls" / "tests"))

import _github_stub  # noqa: E402

REPO = "DysTelefonica/team-skills"
DAY = "2026-11-13"
LABEL = "prerelease:night-red"
TITLE = f"Sin deploy a production exitoso el {DAY}"
DEP_TODAY = {"id": 7, "environment": "production", "created_at": f"{DAY}T08:00:00Z"}
DEP_YESTERDAY = {"id": 5, "environment": "production", "created_at": "2026-11-12T08:00:00Z"}
DEP_STAGING = {"id": 9, "environment": "staging", "created_at": f"{DAY}T09:00:00Z"}


class CheckDailyDeployTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.state_path = self.dir / "state.json"
        self.bin = self.dir / "bin"  # PATH sin gh: el runner del consumer no lo trae
        self.bin.mkdir()

    def tearDown(self):
        self._tmp.cleanup()

    def run_script(self, state=None, *extra):
        """Corre el script contra el stub; devuelve (proceso, estado guardado)."""
        full = {"calls": [], "created_issues": [], **(state or {})}
        with _github_stub.github_stub(self.state_path, full) as base:
            env = _github_stub.stub_env(base)
            env["PATH"] = str(self.bin)
            args = [sys.executable, str(SCRIPT), "--repo", REPO, "--date", DAY] + list(extra)
            proc = subprocess.run(args, capture_output=True, text=True, encoding="utf-8", env=env)
        saved = json.loads(self.state_path.read_text(encoding="utf-8"))
        return proc, saved

    @staticmethod
    def creates(saved):
        return saved["created_issues"]

    # --- sin gh en el PATH: el resultado real, no un doble de gh ---------------

    def test_deploy_check_works_with_no_gh_cli_on_path(self):
        # #1691: el runner del consumer no trae `gh`; el script habla con la API.
        proc, saved = self.run_script({"deployments": [DEP_TODAY],
                                       "deployment_statuses": {"7": [{"state": "success"}]}})
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn(f"ok: deploy a production exitoso el {DAY} (deployment 7)", proc.stdout)
        self.assertEqual(self.creates(saved), [])

    # --- deploy correcto hoy -------------------------------------------------

    def test_successful_production_deploy_today_passes_and_opens_nothing(self):
        proc, saved = self.run_script({"deployments": [DEP_TODAY],
                                       "deployment_statuses": {"7": [{"state": "success"}]}})
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(saved), [])

    def test_latest_status_wins_over_an_earlier_success(self):
        # La API devuelve el estado más reciente primero: `failure` tapa el `success` anterior.
        proc, saved = self.run_script({"deployments": [DEP_TODAY],
                                       "deployment_statuses": {"7": [{"state": "failure"},
                                                                      {"state": "success"}]}})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(len(self.creates(saved)), 1)

    # --- sin deploy correcto hoy: una issue de máxima prioridad ------------

    def test_no_deploy_today_opens_one_top_priority_issue_with_the_fixed_label(self):
        proc, saved = self.run_script({"deployments": [DEP_YESTERDAY]})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        created = self.creates(saved)
        self.assertEqual(len(created), 1)
        self.assertIn(LABEL, created[0]["labels"])
        self.assertIn(DAY, created[0]["title"])
        self.assertIn("máxima", created[0]["body"])

    def test_failed_deploy_today_opens_the_issue(self):
        proc, saved = self.run_script({"deployments": [DEP_TODAY],
                                       "deployment_statuses": {"7": [{"state": "failure"}]}})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(len(self.creates(saved)), 1)

    def test_staging_deploy_does_not_count_as_production(self):
        proc, saved = self.run_script({"deployments": [DEP_STAGING],
                                       "deployment_statuses": {"9": [{"state": "success"}]}})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(len(self.creates(saved)), 1)

    def test_open_issue_with_same_title_is_not_duplicated(self):
        proc, saved = self.run_script({"deployments": [], "issue_list": [{"title": TITLE}]})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(saved), [])

    # --- parámetros ------------------------------------------------------------

    def test_environment_override_is_sent_to_the_api(self):
        proc, saved = self.run_script({"deployments": []}, "--environment", "staging")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        queries = saved["deployment_queries"]
        self.assertTrue(queries, saved)
        self.assertEqual(queries[0]["environment"], ["staging"])

    def test_label_override_is_used_for_the_issue(self):
        proc, saved = self.run_script({"deployments": []}, "--label", "otra:etiqueta")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        created = self.creates(saved)
        self.assertEqual(len(created), 1)
        self.assertIn("otra:etiqueta", created[0]["labels"])

    def test_dry_run_reports_the_issue_and_creates_nothing(self):
        proc, saved = self.run_script({"deployments": []}, "--dry-run")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(saved), [])
        self.assertIn("dry-run", proc.stdout)

    # --- fallo cerrado -----------------------------------------------------------

    def test_unreadable_deployments_fail_closed_without_opening_an_issue(self):
        proc, saved = self.run_script({"deployments_error": True})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(saved), [])

    def test_unreadable_status_fails_closed_without_opening_an_issue(self):
        proc, saved = self.run_script({"deployments": [DEP_TODAY], "statuses_error": True})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(saved), [])

    # --- revision: fecha real, JSON, listado, deduplicacion --------------------

    def test_impossible_date_is_refused_without_opening_an_issue(self):
        proc, saved = self.run_script({"deployments": []}, "--date", "2026-13-45")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(saved), [])
        self.assertNotIn("Traceback", proc.stderr)

    def test_non_json_deployments_fail_closed_without_traceback(self):
        proc, saved = self.run_script({"deployments_raw": "<html>oops</html>"})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(saved), [])
        self.assertNotIn("Traceback", proc.stderr)

    def test_non_json_issue_list_fails_closed_without_traceback(self):
        proc, saved = self.run_script({"deployments": [], "issue_list_raw": "<html>oops</html>"})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(saved), [])
        self.assertNotIn("Traceback", proc.stderr)

    def test_unreadable_issue_list_fails_closed_with_zero_creates(self):
        proc, saved = self.run_script({"deployments": [], "issue_list_error": True})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(saved), [])

    def test_issue_list_is_filtered_by_label_and_has_an_explicit_limit(self):
        proc, saved = self.run_script({"deployments": []})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        queries = saved["issue_list_queries"]
        self.assertEqual(len(queries), 1, saved)
        self.assertEqual(queries[0]["labels"], [LABEL])
        self.assertEqual(queries[0]["state"], ["open"])
        self.assertEqual(queries[0]["per_page"], ["100"])
        self.assertEqual(queries[0]["page"], ["1"])

    def test_issue_creation_failure_fails_closed_without_creating_an_issue(self):
        # El POST que crea la issue no se reintenta: un 5xx no duplica la issue.
        proc, saved = self.run_script({"deployments": [], "issue_post_error": True})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail-closed", proc.stderr)
        self.assertEqual(self.creates(saved), [])
        self.assertEqual([c for c in saved["calls"] if c.startswith("POST")], [
            f"POST repos/{REPO}/issues"])

    def test_open_issue_with_the_same_label_but_another_title_still_opens_one(self):
        proc, saved = self.run_script({"deployments": [], "issue_list": [{"title": "Otra cosa abierta"}]})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(len(self.creates(saved)), 1)


if __name__ == "__main__":
    unittest.main()
