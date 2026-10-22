#!/usr/bin/env python3
# ci-pattern asset tests — changelog fragment gate (HR-44, DysTelefonica/team-skills#665)
"""Black-box suite for the changelog fragment gate.

Each test builds a throwaway git repository, commits a base, applies one
change on top and runs the gate against that pair. The violating fragments
come from ``tests/fixtures/violador/`` on disk, not from inline strings, so
the fixture is the single source of what the gate must reject.

Exit codes under test: 0 clean · 1 findings · 2 invalid input.
"""
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_changelog_fragment.py"
FIXTURE = ASSET / "tests" / "fixtures" / "violador" / "changelog.d"

SKILL_MD = """---
name: {name}
description: "Trigger: test skill."
metadata:
  author: test
  version: "{version}"
---

# {name}
"""


def git(cwd, *args):
    return subprocess.run(["git", "-C", str(cwd), *args],
                          capture_output=True, text=True, check=True)


class Repo:
    def __init__(self, d):
        self.d = Path(d)
        git(self.d, "init", "-q", "-b", "main")
        git(self.d, "config", "user.name", "test")
        git(self.d, "config", "user.email", "test@test")

    def write(self, rel, content):
        p = self.d / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")

    def commit(self, msg):
        git(self.d, "add", "-A")
        git(self.d, "commit", "-q", "-m", msg)
        return git(self.d, "rev-parse", "HEAD").stdout.strip()


def run_gate(repo, base, head="HEAD"):
    r = subprocess.run([sys.executable, str(GATE), "--repo", str(repo),
                        "--base", base, "--head", head],
                       capture_output=True, text=True)
    return r.returncode, r.stdout + r.stderr


class ChangelogFragmentGateTests(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.repo = Repo(self.tmp)
        self.skill = "personal/x/sk"
        self.repo.write(f"{self.skill}/SKILL.md",
                        SKILL_MD.format(name="sk", version="1.0.0"))
        self.base = self.repo.commit("base")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def edit_skill(self):
        self.repo.write(f"{self.skill}/SKILL.md",
                        SKILL_MD.format(name="sk", version="1.0.0") + "\nMas.\n")

    def test_edit_with_valid_patch_fragment_is_clean(self):
        self.edit_skill()
        self.repo.write(f"{self.skill}/changelog.d/fix-a.md",
                        "type: patch\n\nCorrige la regla A.\n")
        self.repo.commit("fix")
        code, out = run_gate(self.tmp, self.base)
        self.assertEqual(code, 0, out)
        self.assertIn("CHANGELOG-FRAGMENT OK", out)

    def test_edit_with_valid_minor_fragment_is_clean(self):
        self.edit_skill()
        self.repo.write(f"{self.skill}/changelog.d/feat-b.md",
                        "type: minor\n\nAñade la regla B.\n")
        self.repo.commit("feat")
        code, out = run_gate(self.tmp, self.base)
        self.assertEqual(code, 0, out)

    def test_skill_edit_without_fragment_fails_and_names_the_skill(self):
        self.edit_skill()
        self.repo.commit("edit sin fragmento")
        code, out = run_gate(self.tmp, self.base)
        self.assertEqual(code, 1, out)
        self.assertIn("changed without an added fragment", out)
        self.assertIn(self.skill, out)

    def test_fragment_with_unknown_type_fails(self):
        # Fixture violador: 'type: major' no es patch ni minor.
        self.edit_skill()
        self.repo.write(f"{self.skill}/changelog.d/bad-type.md",
                        (FIXTURE / "bad-type.md").read_text(encoding="utf-8"))
        self.repo.commit("fragmento con tipo inválido")
        code, out = run_gate(self.tmp, self.base)
        self.assertEqual(code, 1, out)
        self.assertIn("invalid fragment", out)
        self.assertIn("'type: patch' or 'type: minor'", out)

    def test_fragment_without_entry_fails(self):
        # Fixture violador: cabecera válida y ninguna entrada debajo.
        self.edit_skill()
        self.repo.write(f"{self.skill}/changelog.d/no-body.md",
                        (FIXTURE / "no-body.md").read_text(encoding="utf-8"))
        self.repo.commit("fragmento sin entrada")
        code, out = run_gate(self.tmp, self.base)
        self.assertEqual(code, 1, out)
        self.assertIn("no entry after the type line", out)

    def test_fragment_without_type_header_fails(self):
        # Fixture violador: la primera línea no es la cabecera de tipo.
        self.edit_skill()
        self.repo.write(f"{self.skill}/changelog.d/no-header.md",
                        (FIXTURE / "no-header.md").read_text(encoding="utf-8"))
        self.repo.commit("fragmento sin cabecera")
        code, out = run_gate(self.tmp, self.base)
        self.assertEqual(code, 1, out)
        self.assertIn("first line must be", out)

    def test_modified_preexisting_fragment_does_not_satisfy_the_rule(self):
        # Solo cuentan los fragmentos AÑADIDOS por el diff.
        self.repo.write(f"{self.skill}/changelog.d/old.md",
                        "type: patch\n\nEntrada previa.\n")
        base = self.repo.commit("fragmento previo en base")
        self.edit_skill()
        self.repo.write(f"{self.skill}/changelog.d/old.md",
                        "type: patch\n\nEntrada previa, retocada.\n")
        self.repo.commit("retoca fragmento previo")
        code, out = run_gate(self.tmp, base)
        self.assertEqual(code, 1, out)
        self.assertIn("changed without an added fragment", out)

    def test_readme_inside_changelog_d_is_not_a_fragment(self):
        self.edit_skill()
        self.repo.write(f"{self.skill}/changelog.d/README.md",
                        "Cómo escribir un fragmento.\n")
        self.repo.commit("solo README")
        code, out = run_gate(self.tmp, self.base)
        self.assertEqual(code, 1, out)
        self.assertIn("changed without an added fragment", out)

    def test_no_skill_files_changed_is_clean(self):
        self.repo.write("docs/otra.md", "texto\n")
        self.repo.commit("docs fuera de skill")
        code, out = run_gate(self.tmp, self.base)
        self.assertEqual(code, 0, out)

    def test_new_skill_owes_no_fragment(self):
        self.repo.write("personal/x/nueva/SKILL.md",
                        SKILL_MD.format(name="nueva", version="0.1.0"))
        self.repo.commit("skill nueva")
        code, out = run_gate(self.tmp, self.base)
        self.assertEqual(code, 0, out)

    # --- release PR (HR-44, consolidation): recognised only by its exact shape ---

    def release_base(self):
        """A base with two pending fragments and a root CHANGELOG.md."""
        self.repo.write(f"{self.skill}/changelog.d/a.md", "type: patch\n\nRegla A.\n")
        self.repo.write(f"{self.skill}/changelog.d/b.md", "type: minor\n\nRegla B.\n")
        self.repo.write("CHANGELOG.md", "# Changelog\n\n## [Unreleased]\n")
        return self.repo.commit("base con fragmentos")

    def consolidate(self, version="1.1.0", changelog=True, delete=("a", "b")):
        for name in delete:
            (self.repo.d / self.skill / "changelog.d" / f"{name}.md").unlink()
        self.repo.write(f"{self.skill}/SKILL.md",
                        SKILL_MD.format(name="sk", version=version))
        if changelog:
            self.repo.write("CHANGELOG.md", "# Changelog\n\n## [Unreleased]\n\n- v1.1.0\n")

    def test_release_shape_passes_without_an_added_fragment(self):
        base = self.release_base()
        self.consolidate()
        self.repo.commit("release")
        code, out = run_gate(self.tmp, base)
        self.assertEqual(code, 0, out)
        self.assertIn("release", out)

    def test_release_that_also_touches_another_skill_md_line_fails(self):
        base = self.release_base()
        self.consolidate()
        self.repo.write(f"{self.skill}/SKILL.md",
                        SKILL_MD.format(name="sk", version="1.1.0") + "\nMas.\n")
        self.repo.commit("release con otra línea")
        code, out = run_gate(self.tmp, base)
        self.assertEqual(code, 1, out)
        self.assertIn("changed without an added fragment", out)

    def test_release_that_also_touches_an_asset_fails(self):
        base = self.release_base()
        self.consolidate()
        self.repo.write(f"{self.skill}/assets/gate.py", "print('x')\n")
        self.repo.commit("release con asset")
        code, out = run_gate(self.tmp, base)
        self.assertEqual(code, 1, out)
        self.assertIn("changed without an added fragment", out)

    def test_release_without_changelog_change_fails(self):
        base = self.release_base()
        self.consolidate(changelog=False)
        self.repo.commit("release sin changelog")
        code, out = run_gate(self.tmp, base)
        self.assertEqual(code, 1, out)

    def test_version_bump_without_deleting_a_fragment_fails(self):
        base = self.release_base()
        self.consolidate(delete=())
        self.repo.commit("bump a mano")
        code, out = run_gate(self.tmp, base)
        self.assertEqual(code, 1, out)

    def test_deleting_fragments_without_a_version_bump_fails(self):
        base = self.release_base()
        for name in ("a", "b"):
            (self.repo.d / self.skill / "changelog.d" / f"{name}.md").unlink()
        self.repo.write("CHANGELOG.md", "# Changelog\n\n- x\n")
        self.repo.commit("borra fragmentos sin versión")
        code, out = run_gate(self.tmp, base)
        self.assertEqual(code, 1, out)

    def test_unresolvable_base_exits_2(self):
        self.edit_skill()
        self.repo.commit("edit")
        code, out = run_gate(self.tmp, "ref-que-no-existe")
        self.assertEqual(code, 2, out)
        self.assertIn("does not resolve", out)


if __name__ == "__main__":
    unittest.main()
