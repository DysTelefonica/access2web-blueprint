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


class MergeBaseDiff(unittest.TestCase):
    """#T6a (Refs #509): el diff se calcula contra el merge-base, no contra
    la punta de la base del PR. Caso de la auditoría (114818713399,
    115014106789, 115018249635, 115018311601, 115023387817, 115047671163):
    main avanza entre la bifurcación y la apertura del PR. Con dos puntos
    (`base..head`) el commit que entró a main se ve como cambio del PR y
    dispara HR-44; con tres puntos (`merge-base..head`) sólo cuentan los
    commits del PR."""

    def _make_pr_with_main_advance(self, d, *, main_extra_files=(), pr_files=()):
        r = Repo(d)
        r.write_skill("skills/test-skill", "test-skill", "0.3")
        r.commit("v0.3")
        # Bifurcación: la rama de feature parte de aquí.
        git(r.d, "checkout", "-q", "-b", "feature")
        # main avanza después de la bifurcación, con un cambio a la skill
        # que NO bumpa la versión (escenario típico de la auditoría).
        git(r.d, "checkout", "-q", "main")
        skill = r.d / "skills" / "test-skill"
        for rel, content in main_extra_files:
            (skill / rel).write_text(content, encoding="utf-8")
        r.commit("main: extra file without bump")
        # PR: commit en la rama con un cambio a la skill que SÍ bumpa.
        git(r.d, "checkout", "-q", "feature")
        for rel, content in pr_files:
            (skill / rel).write_text(content, encoding="utf-8")
        r.write_skill("skills/test-skill", "test-skill", "0.4")
        r.commit("pr: bump to 0.4")
        return r

    def test_main_extra_commit_does_not_trigger_hr44(self):
        """El commit de main no cuenta como cambio del PR: HR-44 limpio.

        Con diff de 2 puntos (comportamiento anterior) el archivo de main
        aparecía como cambio del PR y, sumado al bump, el gate fallaba
        porque la versión 0.3 de la base y 0.4 de la cabeza se comparaban
        sin bump intermedio (el bump del PR quedaba enmascarado por la
        ausencia de bump en main). Con merge-base el diff sólo contiene
        los commits del PR, la versión sube a 0.4 y el gate es limpio."""
        with tempfile.TemporaryDirectory() as d:
            r = self._make_pr_with_main_advance(
                d,
                main_extra_files=[("main-only.md", "added on main")],
                pr_files=[("pr.md", "added on PR")],
            )
            code, out = run_gate(
                r.d, "main", "feature", "skills/test-skill",
            )
            self.assertEqual((code, "HR-44" in out), (0, False), out)

    def test_pr_with_no_bump_on_top_of_main_advance_is_a_finding(self):
        """El PR introduce un cambio a la skill sin bump: HR-44 rojo,
        y el finding nombra sólo los ficheros del PR (no los de main)."""
        with tempfile.TemporaryDirectory() as d:
            r = self._make_pr_with_main_advance(
                d,
                main_extra_files=[("main-only.md", "added on main")],
                pr_files=[("pr.md", "added on PR")],
            )
            # Sustituir el bump del PR por un cambio sin bump: la versión
            # queda en 0.3 en la cabeza de feature.
            skill = r.d / "skills" / "test-skill"
            (skill / "SKILL.md").write_text(
                SKILL_MD.format(name="test-skill", version="0.3"),
                encoding="utf-8",
            )
            git(r.d, "add", "-A")
            git(r.d, "commit", "-q", "-m", "pr: edit without bump")
            code, out = run_gate(
                r.d, "main", "feature", "skills/test-skill",
            )
            self.assertEqual(code, 1, out)
            self.assertIn("HR-44", out)
            self.assertIn("skills/test-skill", out)
            self.assertNotIn("main-only.md", out)


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
