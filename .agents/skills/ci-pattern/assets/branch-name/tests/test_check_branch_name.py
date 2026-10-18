#!/usr/bin/env python3
"""Suite del gate de nombre de rama (#806 C4), caja negra como en CI.

Salidas: 0 conforme · 1 la rama no cumple · 2 uso o entrada inválidos (regex
ilegible, rama vacía): un gate que no pudo medir no dice OK.
"""
import subprocess
import sys
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_branch_name.py"
P01 = "^(feat|fix|chore|docs|style|refactor|perf|test|build|ci|revert)/[a-z0-9._-]+$"


def run(*args: str):
    out = subprocess.run([sys.executable, str(GATE), *args], capture_output=True, text=True)
    return out.returncode, out.stdout + out.stderr


class TestDefaultConvention(unittest.TestCase):
    def test_conforming_branches_pass(self):
        for ref in ("feat/branch-name-gate", "fix/a", "docs/v1.2_notes", "revert/x-1"):
            with self.subTest(ref=ref):
                self.assertEqual(run(ref)[0], 0, ref)

    def test_non_conforming_branches_fail_and_say_why(self):
        for ref in ("Feat/x", "feature/x", "feat/", "feat/Upper", "feat/with space", "main2",
                    "feat/x/y", "wip", "fix"):
            with self.subTest(ref=ref):
                code, out = run(ref)
                self.assertEqual(code, 1, out)
                self.assertIn(ref, out)
                self.assertIn("<type>/<slug>", out)

    def test_default_pattern_is_p01(self):
        self.assertEqual(run("--pattern", P01, "feat/x")[0], 0)
        self.assertEqual(run("--pattern", P01, "nope")[0], 1)


class TestConsumerPattern(unittest.TestCase):
    def test_pattern_argument_replaces_the_convention(self):
        self.assertEqual(run("--pattern", r"^issue-\d+$", "issue-12")[0], 0)
        self.assertEqual(run("--pattern", r"^issue-\d+$", "feat/x")[0], 1)

    def test_the_whole_name_must_match(self):
        self.assertEqual(run("--pattern", r"feat/[a-z]+", "feat/abc/extra")[0], 1)

    def test_unreadable_pattern_is_exit_two(self):
        code, out = run("--pattern", "(", "feat/x")
        self.assertEqual(code, 2, out)
        self.assertIn("pattern", out)

    def test_empty_ref_is_exit_two(self):
        self.assertEqual(run("")[0], 2)

    def test_missing_ref_is_exit_two(self):
        self.assertEqual(run()[0], 2)


class TestExemptions(unittest.TestCase):
    def test_default_branch_and_archive_are_exempt(self):
        for ref in ("main", "archive/old-thing"):
            with self.subTest(ref=ref):
                self.assertEqual(run(ref)[0], 0, ref)

    def test_allow_lists_grandfathered_branches(self):
        self.assertEqual(run("legacy-branch")[0], 1)
        code, out = run("--allow", "legacy-branch", "legacy-branch")
        self.assertEqual(code, 0, out)
        self.assertIn("grandfathered", out)

    def test_skill_fleet_propagation_branches_are_exempt(self):
        self.assertEqual(run("skill-fleet/APAP_WEB")[0], 0)

    def test_dependabot_branch_is_exempt_only_for_the_bot(self):
        ref = "dependabot/pip/ruff-0.15.0"
        self.assertEqual(run("--actor", "dependabot[bot]", ref)[0], 0)
        self.assertEqual(run("--actor", "someone", ref)[0], 1)
        self.assertEqual(run(ref)[0], 1)

    def test_dependabot_pattern_is_narrow(self):
        self.assertEqual(run("--actor", "dependabot[bot]", "dependabot/evil/x")[0], 1)


if __name__ == "__main__":
    unittest.main()
