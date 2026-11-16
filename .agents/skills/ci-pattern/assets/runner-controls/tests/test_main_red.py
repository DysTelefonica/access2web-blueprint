#!/usr/bin/env python3
"""Tests de main_red.py (#1708, parte A) contra el stub propio `_main_red_stub.py`, sin `gh`."""

import os
import subprocess
import sys
import unittest
from pathlib import Path

ASSET_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from _main_red_stub import main_red_stub  # noqa: E402

ASSET = ASSET_DIR / "main_red.py"
REPO = "owner/repo"
RUN = 4242
SHA = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
RUN_URL = "https://github.com/owner/repo/actions/runs/4242"
LABEL = "prerelease:main-red"


def run(conclusion="failure", branch="main", status="completed"):
    return {"id": RUN, "name": "CI", "head_branch": branch, "head_sha": SHA, "status": status,
            "conclusion": conclusion, "html_url": RUN_URL, "event": "workflow_run"}


def sync(state, *extra):
    """Ejecuta `main_red.py sync` contra el stub; devuelve (proceso, escrituras)."""
    with main_red_stub(state) as base:
        env = {**os.environ, "GITHUB_API_URL": base, "GITHUB_TOKEN": "test-token",
               "GH_API_RETRY_ESPERA_S": "0"}
        env.pop("CI_PATTERN_GH", None)
        proc = subprocess.run([sys.executable, str(ASSET), "sync", "--repo", REPO, "--run", str(RUN), *extra],
                              env=env, capture_output=True, text=True, encoding="utf-8", timeout=60)
    return proc, state["writes"]


def writes_to(writes, method, suffix):
    return [w for w in writes if w["method"] == method and w["path"].endswith(suffix)]


class MainRedSyncTests(unittest.TestCase):
    def test_a_red_default_branch_run_opens_one_issue(self):
        proc, writes = sync({"run": run("failure"), "pulls": [{"number": 7}], "issues": []})
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        opened = writes_to(writes, "POST", f"repos/{REPO}/issues")
        self.assertEqual(len(opened), 1, writes)
        body = opened[0]["body"]
        self.assertTrue(body["title"].startswith("main en rojo: CI 1f4e9a2"), body["title"])
        self.assertEqual(body["labels"], [LABEL])
        self.assertIn(RUN_URL, body["body"])
        self.assertIn("#7", body["body"])
        self.assertIn(SHA, body["body"])

    def test_a_second_red_reuses_the_open_issue_and_comments_the_run(self):
        proc, writes = sync({"run": run("failure"), "pulls": [{"number": 7}],
                             "issues": [{"number": 5}]})
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(writes_to(writes, "POST", f"repos/{REPO}/issues"), [], writes)
        comments = writes_to(writes, "POST", f"repos/{REPO}/issues/5/comments")
        self.assertEqual(len(comments), 1, writes)
        self.assertIn(RUN_URL, comments[0]["body"]["body"])

    def test_a_green_run_closes_every_open_issue_with_a_comment(self):
        proc, writes = sync({"run": run("success"), "issues": [{"number": 5}, {"number": 6}]})
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        for number in (5, 6):
            self.assertEqual(len(writes_to(writes, "POST", f"repos/{REPO}/issues/{number}/comments")), 1, writes)
            closed = writes_to(writes, "PATCH", f"repos/{REPO}/issues/{number}")
            self.assertEqual(closed[0]["body"], {"state": "closed", "state_reason": "completed"}, writes)

    def test_a_green_run_with_none_open_does_nothing(self):
        proc, writes = sync({"run": run("success"), "issues": []})
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(writes, [])

    def test_a_run_off_main_or_cancelled_skips_without_writes(self):
        for name, current in (("other branch", run("failure", branch="feat/x")), ("cancelled", run("cancelled"))):
            with self.subTest(name):
                proc, writes = sync({"run": current, "issues": [{"number": 5}]})
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
                self.assertIn("SKIP", proc.stdout)
                self.assertEqual(writes, [])

    def test_an_unreadable_run_exits_2_and_writes_nothing(self):
        proc, writes = sync({"run": run("failure"), "run_error": True, "issues": []})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(writes, [])
        self.assertIn("sync --repo", proc.stdout)

    def test_a_dry_run_writes_nothing(self):
        proc, writes = sync({"run": run("failure"), "pulls": [{"number": 7}], "issues": []}, "--dry-run")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("DRY-RUN", proc.stdout)
        self.assertEqual(writes, [])

    def test_a_red_without_a_merged_pr_names_it_desconocido(self):
        proc, writes = sync({"run": run("failure"), "pulls": [], "issues": []})
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        opened = writes_to(writes, "POST", f"repos/{REPO}/issues")
        self.assertIn("desconocido", opened[0]["body"]["body"])

    def test_a_failed_issue_write_exits_2_with_the_rerun_command(self):
        proc, _ = sync({"run": run("failure"), "pulls": [{"number": 7}], "issues": [], "write_error": True})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("sync --repo", proc.stdout)


if __name__ == "__main__":
    unittest.main()
