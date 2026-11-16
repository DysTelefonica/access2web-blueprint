#!/usr/bin/env python3
# ci-pattern asset tests — apertura de eslabones de cadena (#1627 slice 7, regla 5)
"""Suite de open_chain_link.py: cada eslabón intermedio lleva `chain:partial` y
`Refs #N` en la primera línea, sin palabra de cierre; solo la punta lleva
`Closes #N`; no se abren más de 3 eslabones abiertos; el modo de ensayo no crea
nada. Ejecución: `python3 test_open_chain_link.py`. Solo stdlib; gh es un doble."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parents[1]
SCRIPT = ASSET / "open_chain_link.py"
FAKE = Path(__file__).resolve().parent / "fake_gh.py"

LIST_EMPTY = {"match": "pr list", "stdout": "[]"}
LIST_TWO = {"match": "pr list", "stdout": json.dumps([{"number": 11}, {"number": 12}])}
LIST_THREE = {"match": "pr list", "stdout": json.dumps([{"number": 11}, {"number": 12}, {"number": 13}])}
CREATE_OK = {"match": "pr create", "stdout": "https://github.com/o/r/pull/99\n"}
HEAD_EXISTS = {"match": "pr list --repo DysTelefonica/team-skills --head",
               "stdout": json.dumps([{"number": 99, "url": "https://github.com/o/r/pull/99"}])}
HEAD_NONE = {"match": "pr list --repo DysTelefonica/team-skills --head", "stdout": "[]"}


class OpenChainLinkTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.log = self.dir / "gh.log"
        self.body = self.dir / "body.md"
        self.body.write_text("Cambio de la cadena.\n", encoding="utf-8")

    def tearDown(self):
        self._tmp.cleanup()

    def run_script(self, rules, *extra):
        if not any("--head" in rule["match"] for rule in rules):
            rules = [HEAD_NONE, *rules]
        rules_path = self.dir / "rules.json"
        rules_path.write_text(json.dumps(rules), encoding="utf-8")
        if self.log.exists():
            self.log.unlink()
        env = dict(os.environ)
        env["PRERELEASE_GH_CMD"] = json.dumps([sys.executable, str(FAKE)])
        env["FAKE_GH_RULES"] = str(rules_path)
        env["FAKE_GH_LOG"] = str(self.log)
        args = [
            sys.executable, str(SCRIPT),
            "--repo", "DysTelefonica/team-skills",
            "--head", "feat/x-link-2",
            "--title", "eslabon 2",
            "--body-file", str(self.body),
        ] + list(extra)
        proc = subprocess.run(args, capture_output=True, text=True, encoding="utf-8", env=env)
        calls = []
        if self.log.exists():
            calls = [json.loads(x) for x in self.log.read_text(encoding="utf-8").splitlines()]
        return proc, calls

    def creates(self, calls):
        return [c for c in calls if c[:2] == ["pr", "create"]]

    # --- eslabón intermedio ---------------------------------------------------

    def test_intermediate_link_is_opened_with_label_and_refs_only(self):
        proc, calls = self.run_script([LIST_EMPTY, CREATE_OK], "--issue", "1627")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        created = self.creates(calls)
        self.assertEqual(len(created), 1)
        args = created[0]
        self.assertIn("chain:partial", args)
        body = args[args.index("--body") + 1]
        self.assertEqual(body.splitlines()[0], "Refs #1627")
        self.assertNotRegex(body, r"(?i)\b(closes|fixes|resolves)\b")

    def test_intermediate_body_with_closing_keyword_is_refused(self):
        self.body.write_text("Closes #1627\n", encoding="utf-8")
        proc, calls = self.run_script([LIST_EMPTY, CREATE_OK], "--issue", "1627")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(calls), [])

    def test_fourth_open_link_is_refused(self):
        proc, calls = self.run_script([LIST_THREE, CREATE_OK], "--issue", "1627")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(calls), [])
        self.assertIn("3", proc.stdout + proc.stderr)

    def test_third_open_link_is_still_allowed(self):
        proc, calls = self.run_script([LIST_TWO, CREATE_OK], "--issue", "1627")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertEqual(len(self.creates(calls)), 1)

    # --- punta ----------------------------------------------------------------

    def test_tip_carries_closes_and_no_chain_label(self):
        proc, calls = self.run_script([LIST_THREE, CREATE_OK], "--issue", "1627", "--tip")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        args = self.creates(calls)[0]
        self.assertNotIn("chain:partial", args)
        body = args[args.index("--body") + 1]
        self.assertEqual(body.splitlines()[0], "Closes #1627")

    # --- ensayo y fallo cerrado -----------------------------------------------

    def test_dry_run_prints_the_command_and_creates_nothing(self):
        proc, calls = self.run_script([LIST_EMPTY, CREATE_OK], "--issue", "1627", "--dry-run")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertEqual(self.creates(calls), [])
        self.assertIn("pr create", proc.stdout)
        self.assertIn("Refs #1627", proc.stdout)

    def test_dry_run_still_refuses_a_fourth_link(self):
        proc, calls = self.run_script([LIST_THREE, CREATE_OK], "--issue", "1627", "--dry-run")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(calls), [])

    def test_unreadable_open_link_count_fails_closed(self):
        broken = {"match": "pr list", "stdout": "", "stderr": "HTTP 500\n", "rc": 1}
        proc, calls = self.run_script([broken, CREATE_OK], "--issue", "1627")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(calls), [])

    def test_missing_issue_is_invalid_input(self):
        proc, calls = self.run_script([LIST_EMPTY, CREATE_OK])
        self.assertEqual(proc.returncode, 2)
        self.assertEqual(self.creates(calls), [])

    # --- idempotencia y palabras de cierre (revision) --------------------------

    def test_existing_open_pr_for_the_head_is_reported_and_nothing_is_created(self):
        proc, calls = self.run_script([HEAD_EXISTS, LIST_EMPTY, CREATE_OK], "--issue", "1627")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.creates(calls), [])
        self.assertIn("pull/99", proc.stdout)

    def test_closing_keywords_in_an_intermediate_body_are_refused(self):
        for line in ("Resolves #1", "Fixes #1", "Closed #1", "fix #1", "RESOLVED #1"):
            with self.subTest(line=line):
                self.body.write_text(line + "\n", encoding="utf-8")
                proc, calls = self.run_script([LIST_EMPTY, CREATE_OK], "--issue", "1627")
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertEqual(self.creates(calls), [])

    def test_closing_keyword_in_the_title_is_refused_for_links_and_tip(self):
        for extra in ([], ["--tip"]):
            with self.subTest(tip=bool(extra)):
                proc, calls = self.run_script([LIST_EMPTY, CREATE_OK], "--issue", "1627",
                                              "--title", "Fixes #1627 y mas", *extra)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertEqual(self.creates(calls), [])

    def run_with_gh(self, gh_cmd, *extra):
        env = dict(os.environ)
        env["PRERELEASE_GH_CMD"] = json.dumps(gh_cmd)
        env["FAKE_GH_RULES"] = str(self.dir / "rules.json")
        args = [sys.executable, str(SCRIPT), "--repo", "DysTelefonica/team-skills", "--head", "feat/x-link-2",
                "--title", "eslabon 2", "--body-file", str(self.body), "--issue", "1627", *extra]
        return subprocess.run(args, capture_output=True, text=True, encoding="utf-8", env=env)

    def test_non_json_output_fails_closed_without_traceback(self):
        rules_path = self.dir / "rules.json"
        rules_path.write_text(json.dumps([{"match": "pr list", "stdout": "<html>oops</html>"}]), encoding="utf-8")
        proc = self.run_with_gh([sys.executable, str(FAKE)])
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertNotIn("Traceback", proc.stderr)

    def test_missing_gh_fails_closed_without_traceback(self):
        proc = self.run_with_gh([str(self.dir / "no-such-gh-binary")])
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertNotIn("Traceback", proc.stderr)


if __name__ == "__main__":
    unittest.main()
