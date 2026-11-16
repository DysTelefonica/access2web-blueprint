#!/usr/bin/env python3
"""Tests de build_night_snapshot.py (#1627, slice 4b): gh falso por CI_PATTERN_GH, solo lectura."""

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
ASSET = HERE.parent / "build_night_snapshot.py"
sys.path.insert(0, str(HERE.parents[1] / "tests"))

from _fake_gh import GH_ENV, install_fake_gh, run_guarded  # noqa: E402

REPO = "owner/repo"
WORKFLOW = "tests.yml"
NIGHT_DATE = "2026-11-13"
GOOD_FINISHED = "2026-11-12T04:20:00Z"

STUB = r'''
import json
import sys

STATE = PLACEHOLDER
with open(STATE, encoding="utf-8") as fh:
    state = json.load(fh)
args = sys.argv[1:]
state.setdefault("calls", []).append(args)
with open(STATE, "w", encoding="utf-8") as fh:
    json.dump(state, fh)
if not args or args[0] != "api":
    sys.stderr.write("fake gh: unsupported command\n")
    sys.exit(1)
path = args[1]
if path == "graphql":
    pr = [a.split("=", 1)[1] for a in args if a.startswith("pr=")][0]
    key = "graphql#" + pr
else:
    key = path
if key not in state["api"]:
    sys.stderr.write("gh: Not Found (HTTP 404) " + key + "\n")
    sys.exit(1)
print(json.dumps(state["api"][key]))
'''

RUNS = f"repos/{REPO}/actions/workflows/{WORKFLOW}/runs?branch=main&event=schedule&per_page=1"
GOOD = f"repos/{REPO}/actions/workflows/{WORKFLOW}/runs?branch=main&event=schedule&status=success&per_page=1"
PULLS = f"repos/{REPO}/pulls?state=closed&base=main&sort=updated&direction=desc&per_page=100"
JOBS = f"repos/{REPO}/actions/runs/900/jobs?per_page=100"
DEPLOYS = f"repos/{REPO}/deployments?environment=production&per_page=50"


def night_run(conclusion="failure", run_id=900):
    return {"id": run_id, "conclusion": conclusion, "status": "completed",
            "html_url": f"https://github.com/{REPO}/actions/runs/{run_id}",
            "created_at": f"{NIGHT_DATE}T03:00:00Z", "updated_at": f"{NIGHT_DATE}T03:40:00Z"}


def pr(number, merged_at, issues):
    return {"number": number, "title": f"PR {number}", "merged_at": merged_at,
            "updated_at": merged_at, "merge_commit_sha": f"{number:040x}"[:40],
            "_issues": issues}


def base_state(pulls=None, jobs=None, deployments=None, good=True, night=None):
    api = {
        RUNS: {"workflow_runs": [night or night_run()]},
        GOOD: {"workflow_runs": [{"id": 800, "conclusion": "success",
                                  "html_url": "u", "updated_at": GOOD_FINISHED}] if good else []},
        PULLS: pulls if pulls is not None else [
            pr(12, "2026-11-13T09:00:00Z", [5]),
            pr(11, "2026-11-12T18:00:00Z", []),
            pr(10, "2026-11-11T10:00:00Z", [4]),
        ],
        JOBS: {"jobs": jobs if jobs is not None else [
            {"name": "e2e (full)", "conclusion": "failure", "html_url": "j1"},
            {"name": "lint", "conclusion": "success", "html_url": "j2"},
            {"name": "trivy", "conclusion": None, "html_url": "j3"},
        ]},
        DEPLOYS: deployments if deployments is not None else [],
    }
    for item in api[PULLS]:
        numbers = item.pop("_issues")
        api[f"graphql#{item['number']}"] = {"data": {"repository": {"pullRequest": {
            "closingIssuesReferences": {"nodes": [{"number": n} for n in numbers]}}}}}
    return {"calls": [], "api": api}


class BuildNightSnapshotTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.state_file = self.root / "state.json"

    def run_tool(self, state, *extra):
        self.state_file.write_text(json.dumps(state), encoding="utf-8")
        shim = install_fake_gh(self.root / "bin", STUB.replace("PLACEHOLDER", repr(str(self.state_file))))
        env = {**os.environ, GH_ENV: str(shim)}
        cmd = [sys.executable, str(ASSET), "--repo", REPO, "--workflow", WORKFLOW,
               "--branch", "main", "--environment", "production", *extra]
        return run_guarded(cmd, env=env, fixture_dir=self.root, capture_output=True,
                           text=True, encoding="utf-8", timeout=120)

    def calls(self):
        return json.loads(self.state_file.read_text(encoding="utf-8"))["calls"]

    def snapshot(self, state, *extra):
        out = self.root / "snap.json"
        proc = self.run_tool(state, "--out", str(out), "--date", NIGHT_DATE, *extra)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        return json.loads(out.read_text(encoding="utf-8"))

    def test_snapshot_has_every_key_night_report_requires(self):
        snap = self.snapshot(base_state())
        for key in ("repository", "night", "last_good_night", "prs", "jobs", "deploy"):
            self.assertIn(key, snap)
        self.assertEqual(snap["repository"], REPO)

    def test_only_prs_merged_after_last_good_night_are_listed_with_their_issues(self):
        prs = self.snapshot(base_state())["prs"]
        self.assertEqual(sorted(p["number"] for p in prs), [11, 12])
        by_number = {p["number"]: p for p in prs}
        self.assertEqual(by_number[12]["issues"], [5])
        self.assertEqual(by_number[11]["issues"], [])
        self.assertEqual(by_number[12]["merge_sha"], f"{12:040x}")

    def test_without_a_last_good_night_every_page_pr_is_listed(self):
        snap = self.snapshot(base_state(good=False))
        self.assertIsNone(snap["last_good_night"])
        self.assertEqual([p["number"] for p in snap["prs"]], [12, 11, 10])

    def test_jobs_are_marked_expensive_by_name_and_keep_their_conclusion(self):
        jobs = {j["name"]: j for j in self.snapshot(base_state())["jobs"]}
        self.assertTrue(jobs["e2e (full)"]["expensive"])
        self.assertFalse(jobs["lint"]["expensive"])
        self.assertEqual(jobs["e2e (full)"]["conclusion"], "failure")
        self.assertIsNone(jobs["trivy"]["conclusion"])

    def test_successful_production_deployment_created_that_utc_day_counts(self):
        deployments = [
            {"id": 77, "created_at": f"{NIGHT_DATE}T05:00:00Z", "environment": "production"},
        ]
        state = base_state(deployments=deployments)
        state["api"][f"repos/{REPO}/deployments/77/statuses?per_page=1"] = [
            {"state": "success", "created_at": f"{NIGHT_DATE}T05:02:00Z"}]
        deploy = self.snapshot(state)["deploy"]
        self.assertTrue(deploy["deployed"])
        self.assertEqual(deploy["at"], f"{NIGHT_DATE}T05:02:00Z")

    def test_failed_status_or_another_day_is_not_a_deployment(self):
        cases = {
            "failed status": ({"id": 77, "created_at": f"{NIGHT_DATE}T05:00:00Z"}, "failure"),
            "other day": ({"id": 77, "created_at": "2026-11-12T05:00:00Z"}, "success"),
        }
        for label, (deployment, status) in cases.items():
            with self.subTest(label=label):
                state = base_state(deployments=[deployment])
                state["api"][f"repos/{REPO}/deployments/77/statuses?per_page=1"] = [
                    {"state": status, "created_at": f"{NIGHT_DATE}T05:02:00Z"}]
                self.assertFalse(self.snapshot(state)["deploy"]["deployed"])

    def test_dry_run_prints_the_snapshot_and_writes_no_file(self):
        proc = self.run_tool(base_state(), "--dry-run", "--date", NIGHT_DATE)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("prs", json.loads(proc.stdout))
        self.assertFalse((self.root / "snap.json").exists())

    def test_without_out_or_dry_run_exits_2(self):
        proc = self.run_tool(base_state(), "--date", NIGHT_DATE)
        self.assertEqual(proc.returncode, 2)

    def test_a_failed_github_read_fails_closed_and_writes_nothing(self):
        state = base_state()
        del state["api"][JOBS]
        out = self.root / "snap.json"
        proc = self.run_tool(state, "--out", str(out), "--date", NIGHT_DATE)
        self.assertEqual(proc.returncode, 2)
        self.assertIn("jobs", proc.stderr)
        self.assertFalse(out.exists())

    def test_a_full_pr_page_that_cannot_be_bounded_fails_closed(self):
        pulls = [pr(1000 - i, "2026-11-13T09:00:00Z", []) for i in range(100)]
        proc = self.run_tool(base_state(pulls=pulls), "--dry-run", "--date", NIGHT_DATE)
        self.assertEqual(proc.returncode, 2)
        self.assertIn("100", proc.stderr)

    def test_the_tool_only_reads_from_github(self):
        self.run_tool(base_state(), "--dry-run", "--date", NIGHT_DATE)
        calls = self.calls()
        self.assertTrue(calls)
        for call in calls:
            self.assertEqual(call[0], "api")
            joined = " ".join(call)
            self.assertNotIn("-X", call)
            self.assertNotIn("--method", call)
            self.assertNotIn("POST", joined.upper().replace("POSTED", ""))


if __name__ == "__main__":
    unittest.main()
