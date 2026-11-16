#!/usr/bin/env python3
"""Tests del informe nocturno en HTML (#1627, slice 4). stdlib, sin red ni gh.

El informe se genera desde una instantánea JSON; estos tests fijan el contrato de esa
instantánea, el filtrado desde la última noche buena, el agrupado por issue, el
escapado HTML, la tabla de trabajos caros, el deploy y el modo de ensayo que no escribe.
"""

from __future__ import annotations

import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCRIPT = HERE.parent / "night_report.py"
spec = importlib.util.spec_from_file_location("night_report", SCRIPT)
night = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(night)


def snapshot(**overrides) -> dict:
    base = {
        "repository": "DysTelefonica/team-skills",
        "night": {"sha": "b" * 40, "started_at": "2026-11-13T02:00:00+00:00"},
        "last_good_night": {"sha": "a" * 40, "finished_at": "2026-11-12T04:00:00+00:00"},
        "prs": [
            {"number": 1001, "title": "fix(a): uno", "merged_at": "2026-11-12T10:00:00+00:00",
             "merge_sha": "c" * 40, "issues": [900]},
            {"number": 1002, "title": "feat(b): dos <script>", "merged_at": "2026-11-12T11:00:00+00:00",
             "merge_sha": "d" * 40, "issues": [900, 901]},
            {"number": 1003, "title": "docs: sin issue", "merged_at": "2026-11-12T12:00:00+00:00",
             "merge_sha": "e" * 40, "issues": []},
            {"number": 999, "title": "antes de la noche buena", "merged_at": "2026-11-11T09:00:00+00:00",
             "merge_sha": "f" * 40, "issues": [899]},
        ],
        "jobs": [
            {"name": "full", "expensive": True, "conclusion": "success", "run_url": "https://ci/full"},
            {"name": "nightly-unit", "expensive": True, "conclusion": "failure", "run_url": "https://ci/unit"},
            {"name": "lint", "expensive": False, "conclusion": "success", "run_url": "https://ci/lint"},
        ],
        "deploy": {"deployed": False, "sha": "b" * 40, "at": None},
    }
    base.update(overrides)
    return base


class SelectionTests(unittest.TestCase):
    def test_only_prs_merged_after_the_last_good_night_are_listed(self) -> None:
        numbers = [pr["number"] for pr in night.prs_since_last_good(snapshot())]
        self.assertEqual(numbers, [1001, 1002, 1003])

    def test_without_a_good_night_every_pr_is_listed(self) -> None:
        numbers = [pr["number"] for pr in night.prs_since_last_good(snapshot(last_good_night=None))]
        self.assertEqual(numbers, [999, 1001, 1002, 1003])

    def test_prs_are_grouped_by_linked_issue_and_unlinked_ones_are_kept(self) -> None:
        groups = dict(night.group_by_issue(night.prs_since_last_good(snapshot())))
        self.assertEqual([p["number"] for p in groups[900]], [1001, 1002])
        self.assertEqual([p["number"] for p in groups[901]], [1002])
        self.assertEqual([p["number"] for p in groups[None]], [1003])


class RenderTests(unittest.TestCase):
    def test_each_pr_is_linked_and_its_issues_are_linked(self) -> None:
        html = night.render(snapshot())
        self.assertIn("https://github.com/DysTelefonica/team-skills/pull/1001", html)
        self.assertIn("https://github.com/DysTelefonica/team-skills/issues/900", html)
        self.assertIn("https://github.com/DysTelefonica/team-skills/issues/901", html)
        self.assertNotIn("pull/999", html)

    def test_titles_are_html_escaped(self) -> None:
        html = night.render(snapshot())
        self.assertNotIn("<script>", html)
        self.assertIn("&lt;script&gt;", html)

    def test_expensive_jobs_show_their_conclusion_and_only_expensive_ones_are_listed(self) -> None:
        html = night.render(snapshot())
        self.assertIn("nightly-unit", html)
        self.assertIn("failure", html)
        self.assertIn("https://ci/unit", html)
        self.assertNotIn(">lint<", html)

    def test_an_expensive_job_without_result_says_so(self) -> None:
        html = night.render(snapshot(jobs=[{"name": "full", "expensive": True, "conclusion": None,
                                            "run_url": None}]))
        self.assertIn("sin resultado", html)

    def test_deploy_state_is_reported(self) -> None:
        self.assertIn("no desplegado", night.render(snapshot()))
        deployed = snapshot(deploy={"deployed": True, "sha": "b" * 40, "at": "2026-11-13T05:00:00+00:00"})
        self.assertIn("desplegado", night.render(deployed).replace("no desplegado", ""))

    def test_the_same_snapshot_renders_the_same_html_twice(self) -> None:
        self.assertEqual(night.render(snapshot()), night.render(snapshot()))


class CliTests(unittest.TestCase):
    def run_cli(self, args: list[str]) -> subprocess.CompletedProcess:
        return subprocess.run([sys.executable, str(SCRIPT), *args], capture_output=True, text=True,
                              check=False)

    def test_dry_run_writes_nothing_and_prints_the_html(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            snap = Path(tmp) / "snap.json"
            out = Path(tmp) / "report.html"
            snap.write_text(json.dumps(snapshot()), encoding="utf-8")
            result = self.run_cli(["--snapshot", str(snap), "--out", str(out), "--dry-run"])
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn("<html", result.stdout)
            self.assertFalse(out.exists())
            self.assertEqual(sorted(p.name for p in Path(tmp).iterdir()), ["snap.json"])

    def test_writing_twice_gives_identical_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            snap = Path(tmp) / "snap.json"
            out = Path(tmp) / "report.html"
            snap.write_text(json.dumps(snapshot()), encoding="utf-8")
            self.assertEqual(self.run_cli(["--snapshot", str(snap), "--out", str(out)]).returncode, 0)
            first = out.read_bytes()
            self.assertEqual(self.run_cli(["--snapshot", str(snap), "--out", str(out)]).returncode, 0)
            self.assertEqual(out.read_bytes(), first)

    def test_an_invalid_snapshot_exits_two_naming_the_problem(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            snap = Path(tmp) / "snap.json"
            snap.write_text("{}", encoding="utf-8")
            result = self.run_cli(["--snapshot", str(snap), "--dry-run"])
            self.assertEqual(result.returncode, 2)
            self.assertIn("repository", result.stderr)

    def test_without_out_or_dry_run_it_refuses(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            snap = Path(tmp) / "snap.json"
            snap.write_text(json.dumps(snapshot()), encoding="utf-8")
            self.assertEqual(self.run_cli(["--snapshot", str(snap)]).returncode, 2)


class EmojiTitleTests(unittest.TestCase):
    def test_an_emoji_title_does_not_crash_under_a_non_utf8_console(self) -> None:
        """Un título con emoji no tumba el script con UnicodeEncodeError (#1627)."""
        for title in ("fix(a): \u2705 listo", "feat(b): \U0001f680 deploy"):
            with self.subTest(title=title), tempfile.TemporaryDirectory() as tmp:
                prs = snapshot()["prs"]
                prs[0]["title"] = title
                snap = Path(tmp) / "snap.json"
                snap.write_text(json.dumps(snapshot(prs=prs)), encoding="utf-8")
                env = {**os.environ, "PYTHONIOENCODING": "cp1252", "PYTHONUTF8": "0"}
                result = subprocess.run([sys.executable, str(SCRIPT), "--snapshot", str(snap), "--dry-run"],
                                        capture_output=True, env=env, check=False)
                self.assertEqual(result.returncode, 0, result.stderr.decode("utf-8", "replace"))
                self.assertNotIn(b"UnicodeEncodeError", result.stderr)




if __name__ == "__main__":
    unittest.main()
