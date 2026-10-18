#!/usr/bin/env python3
# ci-pattern asset tests — release gate: annotated semver tag on the default branch (#195)
"""Executable suite for the release gate (black-box, like CI).

Every test builds a throwaway git repository locally (subprocess git, no
network) and runs the gate against it. Rules under test:

- HR-19: the release tag MUST be annotated and MUST match
  ``^v(0|[1-9]\\d*)\\.(0|[1-9]\\d*)\\.(0|[1-9]\\d*)$`` (strict semver, no
  leading zeros), and the version MUST be greater than the last existing
  semver tag.
- HR-20: the tagged commit MUST be an ancestor of the default branch — a
  hotfix never ships from a side branch.

Exit codes: 0 clean · 1 findings · 2 invalid input or empty subject
(missing tag, missing repository, no commits)."""
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_release_tag.py"


def git(cwd, *args):
    return subprocess.run(["git", "-C", str(cwd), *args],
                          capture_output=True, text=True, check=True)


class Repo:
    """Minimal local git repository fixture; no network, all local."""

    def __init__(self, d):
        self.d = Path(d)
        git(self.d, "init", "-q", "-b", "main")
        git(self.d, "config", "user.name", "fixture")
        git(self.d, "config", "user.email", "fixture@example.test")

    def commit(self, msg="c", allow_empty=True):
        args = ["commit", "-q", "--allow-empty", "-m", msg]
        r = subprocess.run(["git", "-C", str(self.d), *args], capture_output=True, text=True)
        if r.returncode != 0:
            (self.d / "f.txt").write_text(msg, encoding="utf-8")
            git(self.d, "add", "-A")
            git(self.d, "commit", "-q", "-m", msg)
        return git(self.d, "rev-parse", "HEAD").stdout.strip()

    def branch(self, name, start=None):
        git(self.d, "switch", "-q", "-c", name, *(["main"] if start is None else [start]))

    def tag_annotated(self, name, msg="release", ref="HEAD"):
        git(self.d, "tag", "-a", name, "-m", msg, ref)

    def tag_light(self, name, ref="HEAD"):
        git(self.d, "tag", name, ref)


class GateCase(unittest.TestCase):
    def run_gate(self, repo, tag, extra=()):
        out = subprocess.run(
            [sys.executable, str(GATE), "--repo", str(repo.d), "--tag", tag,
             "--default-branch", "main", *extra],
            capture_output=True, text=True)
        return out.returncode, out.stdout + out.stderr

    def table(self, cases, repo):
        for label, tag, code, needle in cases:
            with self.subTest(case=label):
                got, out = self.run_gate(repo, tag)
                self.assertEqual(got, code, out)
                if needle:
                    self.assertIn(needle, out)


class CleanRelease(GateCase):
    def test_annotated_semver_on_main_is_clean(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.tag_annotated("v1.0.0")
            self.table([("annotated semver on main", "v1.0.0", 0, "RELEASE OK")], r)

    def test_greater_patch_bump_is_clean(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.tag_annotated("v1.2.3")
            r.commit("fix")
            r.tag_annotated("v1.2.4")
            self.table([("patch bump", "v1.2.4", 0, "RELEASE OK")], r)

    def test_minor_and_major_bumps_are_clean(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.tag_annotated("v1.2.3")
            r.commit("more")
            r.tag_annotated("v2.0.0")
            self.table([("major bump", "v2.0.0", 0, "RELEASE OK")], r)


class HR19TagShape(GateCase):
    def test_findings(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.tag_annotated("v1.2.3")
            r.commit("work")
            r.tag_annotated("v1.2.4")
            r.tag_light("v1.2.5")
            r.tag_annotated("v1.2")
            r.tag_annotated("v01.2.4")
            r.tag_annotated("v1.2.4-rc.1")
            r.tag_annotated("v1.2.2")
            self.table([
                ("lightweight tag", "v1.2.5", 1, "HR-19"),
                ("two components", "v1.2", 1, "HR-19"),
                ("leading zero", "v01.2.4", 1, "HR-19"),
                ("prerelease suffix", "v1.2.4-rc.1", 1, "HR-19"),
                ("version not greater than last tag", "v1.2.2", 1, "HR-19"),
            ], r)

    def test_negatives(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.tag_annotated("v1.2.3")
            r.commit("work")
            r.tag_annotated("v1.2.4")
            self.table([
                ("annotated exact semver", "v1.2.4", 0, None),
                ("version greater than last", "v1.2.4", 0, None),
            ], r)

    def test_degenerate_zero_padded_major_is_a_finding(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.tag_annotated("v0.0.0")
            r.commit("work")
            r.tag_annotated("v0.0.00")
            self.table([("zero-padded patch", "v0.0.00", 1, "HR-19")], r)


class HR20AncestorOfDefaultBranch(GateCase):
    def test_findings(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.branch("hotfix-side")
            r.commit("side fix")
            r.tag_annotated("v1.0.1")
            self.table([
                ("tagged commit not on default branch", "v1.0.1", 1, "HR-20"),
            ], r)

    def test_negative_hotfix_merged_to_main(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.tag_annotated("v1.0.0")
            r.branch("hotfix")
            r.commit("hotfix")
            git(r.d, "switch", "-q", "main")
            git(r.d, "merge", "-q", "--no-ff", "-m", "merge hotfix", "hotfix")
            r.tag_annotated("v1.0.1")
            self.table([("hotfix merged into main", "v1.0.1", 0, None)], r)

    def test_degenerate_side_branch_tag_on_stale_history(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.branch("experimental")
            r.commit("exp 1")
            r.commit("exp 2")
            r.tag_annotated("v9.9.9")
            self.table([("side branch with greater version", "v9.9.9", 1, "HR-20")], r)


class InvalidInput(GateCase):
    def test_missing_tag_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            code, out = self.run_gate(r, "v1.0.0")
            self.assertEqual(code, 2)

    def test_missing_repository_exits_two(self):
        code, out = self.run_gate(type("R", (), {"d": "/no/such/repo"})(), "v1.0.0")
        self.assertEqual(code, 2)

    def test_empty_repository_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)  # no commits
            code, out = self.run_gate(r, "v1.0.0")
            self.assertEqual(code, 2)

    def test_missing_required_argument(self):
        out = subprocess.run([sys.executable, str(GATE)], capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)


class Contract(GateCase):
    def test_findings_aggregate(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.branch("side")
            r.commit("side")
            r.tag_light("not-semver")  # lightweight + bad name + not on main
            code, out = self.run_gate(r, "not-semver")
            self.assertEqual(code, 1)
            self.assertIn("HR-19", out)
            self.assertIn("HR-20", out)

    def test_ok_reports_measurement(self):
        with tempfile.TemporaryDirectory() as d:
            r = Repo(d)
            r.commit("base")
            r.tag_annotated("v1.0.0")
            code, out = self.run_gate(r, "v1.0.0")
            self.assertEqual(code, 0)
            self.assertIn("RELEASE OK", out)


if __name__ == "__main__":
    unittest.main(verbosity=2)
