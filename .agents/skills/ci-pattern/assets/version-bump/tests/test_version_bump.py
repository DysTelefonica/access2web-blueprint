#!/usr/bin/env python3
# ci-pattern asset tests — version bump gate (HR-44, DysTelefonica/team-skills#156)
"""Executable suite for the version bump gate (black-box).

HR-44: every material edit of a skill MUST increment metadata.version;
a mirror sync MUST start from a resolved SHA of the source. The gate takes
a base ref and a head ref, diffs the skill directory, and reports:

- skill files changed between base and head with the same metadata.version
  in SKILL.md frontmatter → finding (HR-44);
- frontmatter unreadable at either ref → exit 2 (fail-closed);
- no skill files changed → clean (0).

Exit codes: 0 clean · 1 findings · 2 invalid input."""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_version_bump.py"


def git(cwd, *args):
    return subprocess.run(["git", "-C", str(cwd), *args],
                          capture_output=True, text=True, check=True)


SKILL_MD = """---
name: {name}
description: "Trigger: test skill."
metadata:
  author: test
  version: "{version}"
---

# {name}
"""


class Repo:
    def __init__(self, d):
        self.d = Path(d)
        git(self.d, "init", "-q", "-b", "main")
        git(self.d, "config", "user.name", "test")
        git(self.d, "config", "user.email", "test@test")

    def write_skill(self, skill_rel, name, version, extra_files=None):
        skill = self.d / skill_rel
        skill.mkdir(parents=True, exist_ok=True)
        (skill / "SKILL.md").write_text(SKILL_MD.format(name=name, version=version),
                                        encoding="utf-8")
        if extra_files:
            for rel, content in extra_files.items():
                (skill / rel).write_text(content, encoding="utf-8")

    def commit(self, msg):
        git(self.d, "add", "-A")
        git(self.d, "commit", "-q", "-m", msg)

    def tag(self, name):
        git(self.d, "tag", name)


def run_gate(repo, base, head, skill_dir, extra=()):
    argv = [sys.executable, str(GATE), "--repo", str(repo),
            "--base", base, "--head", head]
    if skill_dir:
        argv += ["--skill-dir", skill_dir]
    out = subprocess.run([*argv, *extra], capture_output=True, text=True)
    return out.returncode, out.stdout + out.stderr


class Clean(unittest.TestCase):
    def test_version_bumped_is_clean(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.write_skill("skills/test-skill", "test-skill", "0.3")
            r.commit("v0.3")
            r.write_skill("skills/test-skill", "test-skill", "0.4",
                          {"rules.md": "new rule"})
            r.commit("v0.4")
            code, out = run_gate(r.d, "main~1", "main", "skills/test-skill")
            self.assertEqual((code, "HR-44" in out), (0, False))

    def test_no_skill_changes_is_clean(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.write_skill("skills/test-skill", "test-skill", "0.3")
            r.commit("v1")
            (r.d / "other.txt").write_text("unrelated\n", encoding="utf-8")
            r.commit("unrelated")
            code, out = run_gate(r.d, "main~1", "main", "skills/test-skill")
            self.assertEqual((code, "HR-44" in out), (0, False))

    def test_new_skill_is_clean(self):
        """A skill that didn't exist at base has no base version to compare."""
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            (r.d / "README.md").write_text("# repo\n", encoding="utf-8")
            r.commit("init")
            r.write_skill("skills/new-skill", "new-skill", "0.1")
            r.commit("new skill")
            code, out = run_gate(r.d, "main~1", "main", "skills/new-skill")
            self.assertEqual((code, "HR-44" in out), (0, False))


class Findings(unittest.TestCase):
    def test_material_edit_without_bump_is_a_finding(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.write_skill("skills/test-skill", "test-skill", "0.3")
            r.commit("v0.3")
            r.write_skill("skills/test-skill", "test-skill", "0.3",
                          {"rules.md": "new rule without bump"})
            r.commit("edit without bump")
            code, out = run_gate(r.d, "main~1", "main", "skills/test-skill")
            self.assertEqual(code, 1)
            self.assertIn("HR-44", out)

    def test_semver_downgrade_is_a_finding(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.write_skill("skills/test-skill", "test-skill", "0.4")
            r.commit("v0.4")
            r.write_skill("skills/test-skill", "test-skill", "0.3")
            r.commit("downgrade")
            code, out = run_gate(r.d, "main~1", "main", "skills/test-skill")
            self.assertEqual(code, 1)
            self.assertIn("HR-44", out)


class Discovery(unittest.TestCase):
    def test_without_skill_dir_gates_every_touched_skill(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.write_skill("skills/one", "one", "1.0")
            r.write_skill("personal/owner/two", "two", "1.0")
            r.commit("v1")
            (r.d / "other.txt").write_text("unrelated\n", encoding="utf-8")
            r.commit("no skill files")
            code, out = run_gate(r.d, "main~1", "main", None)
            self.assertEqual((code, "HR-44" in out), (0, False))
            r.write_skill("skills/one", "one", "1.0", {"rules.md": "no bump"})
            r.write_skill("personal/owner/two", "two", "1.1")
            r.commit("edit both skills")
            code, out = run_gate(r.d, "main~1", "main", None)
            self.assertEqual(code, 1)
            self.assertIn("skills/one", out)
            self.assertNotIn("owner/two", out)


class InvalidInput(unittest.TestCase):
    def test_unreadable_frontmatter_at_head_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.write_skill("skills/test-skill", "test-skill", "0.3")
            r.commit("v0.3")
            (r.d / "skills/test-skill/SKILL.md").write_text("no frontmatter\n",
                                                            encoding="utf-8")
            r.commit("broken frontmatter")
            code, out = run_gate(r.d, "main~1", "main", "skills/test-skill")
            self.assertEqual(code, 2)

    def test_unreadable_frontmatter_at_base_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            (r.d / "README.md").write_text("# repo\n", encoding="utf-8")
            r.commit("init")
            (r.d / "skills/test-skill").mkdir(parents=True, exist_ok=True)
            (r.d / "skills/test-skill/SKILL.md").write_text("no frontmatter\n",
                                                            encoding="utf-8")
            r.commit("broken base")
            r.write_skill("skills/test-skill", "test-skill", "0.4",
                          {"new.txt": "content"})
            r.commit("head with good frontmatter")
            code, out = run_gate(r.d, "main~1", "main", "skills/test-skill")
            self.assertEqual(code, 2)

    def test_missing_base_ref_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.write_skill("skills/test-skill", "test-skill", "0.3")
            r.commit("only")
            code, out = run_gate(r.d, "nonexistent-ref", "main", "skills/test-skill")
            self.assertEqual(code, 2)

    def test_missing_required_argument_exits_two(self):
        out = subprocess.run([sys.executable, str(GATE)], capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
