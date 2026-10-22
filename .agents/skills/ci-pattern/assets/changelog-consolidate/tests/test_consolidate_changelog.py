#!/usr/bin/env python3
# ci-pattern asset tests — changelog consolidation (HR-44, DysTelefonica/team-skills#665)
"""Black-box suite for the changelog consolidation.

Each test builds a skill directory and a CHANGELOG.md in a temporary folder,
runs the script as a subprocess and checks the result on disk. Exit codes
under test: 0 done or nothing pending · 1 plan only · 2 invalid input.
"""
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "consolidate_changelog.py"

SKILL_MD = """---
name: sk
description: "Trigger: test skill."
metadata:
  author: test
  version: "{version}"
---

# sk
"""

CHANGELOG = """# Changelog

## [Unreleased]

### Changed

- **`otra`** (v1.0.0 → v1.0.1): entrada previa.

## Previous
"""


class ConsolidateTests(unittest.TestCase):

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.skill = self.tmp / "sk"
        (self.skill / "changelog.d").mkdir(parents=True)
        (self.skill / "SKILL.md").write_text(SKILL_MD.format(version="1.19.12"),
                                             encoding="utf-8")
        self.changelog = self.tmp / "CHANGELOG.md"
        self.changelog.write_text(CHANGELOG, encoding="utf-8")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def fragment(self, name, body):
        (self.skill / "changelog.d" / name).write_text(body, encoding="utf-8")

    def run_script(self, *extra):
        r = subprocess.run([sys.executable, str(SCRIPT), "--skill-dir", str(self.skill),
                            "--changelog", str(self.changelog), *extra],
                           capture_output=True, text=True)
        return r.returncode, r.stdout + r.stderr

    def version(self):
        return next(line for line in (self.skill / "SKILL.md").read_text(encoding="utf-8")
                    .splitlines() if line.strip().startswith("version:"))

    def test_patch_only_bumps_patch(self):
        self.fragment("a.md", "type: patch\n\nCorrige la regla A.\n")
        code, out = self.run_script("--write")
        self.assertEqual(code, 0, out)
        self.assertIn('version: "1.19.13"', self.version())

    def test_any_minor_bumps_minor_and_resets_patch(self):
        self.fragment("a.md", "type: patch\n\nCorrige A.\n")
        self.fragment("b.md", "type: minor\n\nAñade B.\n")
        code, out = self.run_script("--write")
        self.assertEqual(code, 0, out)
        self.assertIn('version: "1.20.0"', self.version())

    def test_bullet_lands_under_unreleased_changed_with_versions(self):
        self.fragment("a.md", "type: patch\n\nCorrige A.\n")
        code, out = self.run_script("--write")
        self.assertEqual(code, 0, out)
        text = self.changelog.read_text(encoding="utf-8")
        self.assertIn("- **`sk`** (v1.19.12 → v1.19.13):\n  - Corrige A.\n", text)
        self.assertLess(text.index("- **`sk`**"), text.index("- **`otra`**"))

    def test_one_sub_bullet_per_fragment_under_the_version_header(self):
        self.fragment("a.md", "type: patch\n\nCorrige A.\n")
        self.fragment("b.md", "type: minor\n\nAñade B\ncon dos líneas.\n")
        code, out = self.run_script("--write")
        self.assertEqual(code, 0, out)
        text = self.changelog.read_text(encoding="utf-8")
        self.assertIn("- **`sk`** (v1.19.12 → v1.20.0):\n"
                      "  - Corrige A.\n"
                      "  - Añade B con dos líneas.\n", text)
        self.assertNotIn("Corrige A. Añade B", text)

    def test_consumed_fragments_are_deleted_and_readme_kept(self):
        self.fragment("a.md", "type: patch\n\nCorrige A.\n")
        self.fragment("README.md", "Cómo escribir un fragmento.\n")
        code, out = self.run_script("--write")
        self.assertEqual(code, 0, out)
        left = sorted(p.name for p in (self.skill / "changelog.d").iterdir())
        self.assertEqual(left, ["README.md"])

    def test_second_run_finds_nothing_pending(self):
        self.fragment("a.md", "type: patch\n\nCorrige A.\n")
        self.run_script("--write")
        code, out = self.run_script("--write")
        self.assertEqual(code, 0, out)
        self.assertIn("nada pendiente", out)
        self.assertIn('version: "1.19.13"', self.version())

    def test_without_write_only_prints_the_plan(self):
        self.fragment("a.md", "type: minor\n\nAñade A.\n")
        code, out = self.run_script()
        self.assertEqual(code, 1, out)
        self.assertIn("v1.19.12 → v1.20.0", out)
        self.assertIn('version: "1.19.12"', self.version())
        self.assertTrue((self.skill / "changelog.d" / "a.md").exists())

    def test_invalid_fragment_exits_2_and_changes_nothing(self):
        self.fragment("a.md", "type: patch\n\nCorrige A.\n")
        self.fragment("bad.md", "type: major\n\nNo existe.\n")
        before = (self.skill / "SKILL.md").read_text(encoding="utf-8")
        code, out = self.run_script("--write")
        self.assertEqual(code, 2, out)
        self.assertIn("first line must be", out)
        self.assertEqual((self.skill / "SKILL.md").read_text(encoding="utf-8"), before)
        self.assertTrue((self.skill / "changelog.d" / "a.md").exists())

    def test_changelog_without_unreleased_exits_2(self):
        self.changelog.write_text("# Changelog\n\n## Previous\n", encoding="utf-8")
        self.fragment("a.md", "type: patch\n\nCorrige A.\n")
        code, out = self.run_script("--write")
        self.assertEqual(code, 2, out)
        self.assertIn("[Unreleased]", out)

    def test_missing_version_exits_2(self):
        (self.skill / "SKILL.md").write_text("---\nname: sk\n---\n", encoding="utf-8")
        self.fragment("a.md", "type: patch\n\nCorrige A.\n")
        code, out = self.run_script("--write")
        self.assertEqual(code, 2, out)
        self.assertIn("metadata.version not found", out)


if __name__ == "__main__":
    unittest.main()
