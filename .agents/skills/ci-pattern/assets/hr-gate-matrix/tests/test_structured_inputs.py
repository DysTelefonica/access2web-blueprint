#!/usr/bin/env python3
# ci-pattern asset tests — structured inputs contract gate (HR-1, DysTelefonica/team-skills#245)
"""Executable suite for the structured inputs contract gate (black-box).

HR-1: gates read government data from structured sources (event payloads,
JSON documents, git refs, argv subjects) and never scrape it out of host
text output. The gate reads the HR→gate matrix, takes every declared gate
asset, and reports ``gh pr view`` / ``gh issue view`` / ``gh pr list`` /
``gh issue list`` without ``--json`` (host text output scraping of
title/body/labels). Pattern-validated subjects (conventional commits,
branch names) are the governed datum itself and stay allowed.

Exit codes: 0 clean · 1 findings · 2 invalid input (unreadable matrix,
missing gate file, empty subject — a matrix with no gate entries)."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_structured_inputs.py"


def build_skill(root: Path, gates: dict, with_hr32=True):
    """Temp skill root: matrix + gate files. gates = {relpath: content}."""
    skill = Path(root, "personal", "ci-pattern")
    (skill / "references").mkdir(parents=True)
    rules = []
    for path, content in gates.items():
        p = skill / path
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")
        rules.append({"id": f"HR-{len(rules) + 1}", "enforcement": "gate",
                      "asset": path, "test": path})
    if with_hr32:
        x = skill / "assets" / "x.py"
        x.parent.mkdir(parents=True, exist_ok=True)
        x.write_text("x = 1\n", encoding="utf-8")
        rules.append({"id": "HR-32", "enforcement": "gate",
                      "asset": "assets/x.py", "test": "assets/x.py"})
    matrix = {"rules": rules}
    (skill / "references" / "hr-gate-matrix.json").write_text(
        json.dumps(matrix), encoding="utf-8")
    return skill


def run_gate(skill, extra=()):
    out = subprocess.run([sys.executable, str(GATE), "--skill-root", str(skill), *extra],
                         capture_output=True, text=True)
    return out.returncode, out.stdout + out.stderr


class Clean(unittest.TestCase):
    def test_gates_without_scraping_are_clean(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {
                "assets/a.py": 'doc = json.loads(Path(args.event_file).read_text())\n',
                "assets/b.sh": 'gh pr view 1 --json title,body\n'})
            code, out = run_gate(skill)
            self.assertEqual((code, "HR-1" in out), (0, False))

    def test_pattern_validated_subjects_stay_allowed(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {
                "assets/c.py": 'subject = git_log_line  # validated against conventional pattern\n'})
            code, out = run_gate(skill)
            self.assertEqual(code, 0)

    def test_degenerate_matrix_without_gate_entries_exits_two(self):
        """HR-18/HR-3: un sujeto vacío no se aprueba a sí mismo."""
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {}, with_hr32=False)
            code, out = run_gate(skill)
            self.assertEqual(code, 2)
            self.assertIn("empty subject", out)


class Findings(unittest.TestCase):
    def test_gh_pr_view_without_json(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {
                "assets/a.sh": 'title=$(gh pr view 123)\n'})
            code, out = run_gate(skill)
            self.assertEqual(code, 1)
            self.assertIn("HR-1", out)

    def test_gh_issue_view_without_json(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {
                "assets/a.sh": 'gh issue view 5\n'})
            code, out = run_gate(skill)
            self.assertEqual(code, 1)
            self.assertIn("HR-1", out)

    def test_gh_pr_list_without_json(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {
                "assets/a.py": 'subprocess.run(["gh", "pr", "list"])\n'})
            code, out = run_gate(skill)
            self.assertEqual(code, 1)
            self.assertIn("HR-1", out)

    def test_findings_cite_gate_line_and_hr_ids(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {
                "assets/deep/g.sh": 'gh issue view 9\n'})
            code, out = run_gate(skill)
            self.assertEqual(code, 1)
            self.assertIn("assets/deep/g.sh", out)
            self.assertIn("(HR-1)", out)
            self.assertNotIn("(H, R", out)

    def test_multiline_python_call_does_not_evade(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {
                "assets/a.py": 'subprocess.run([\n    "gh",\n    "pr",\n    "view",\n    "1",\n])\n'})
            code, out = run_gate(skill)
            self.assertEqual(code, 1)
            self.assertIn("HR-1", out)

    def test_multiline_shell_continuation_does_not_evade(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {
                "assets/a.sh": 'gh \\\n  pr view 1 \\\n  --repo example/repo\n'})
            code, out = run_gate(skill)
            self.assertEqual(code, 1)
            self.assertIn("HR-1", out)

    def test_json_flag_on_same_line_stays_clean(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {
                "assets/a.sh": 'gh pr view 123 --json title --jq .title\n'})
            code, out = run_gate(skill)
            self.assertEqual((code, "HR-1" in out), (0, False))


class Dogfooding(unittest.TestCase):
    """HR-1: el gate se ejecuta contra la skill real y no puede saltarse
    en team-skills — cablearlo en real no puede romper main."""

    def is_team_skills(self) -> bool:
        return (ASSET.parents[4] / "testing" / "run-unit-suites.sh").is_file()

    def test_real_skill_is_clean(self):
        if not self.is_team_skills():
            self.skipTest("instalación fuera de team-skills")
        out = subprocess.run([sys.executable, str(GATE)],
                             capture_output=True, text=True)
        self.assertEqual(out.returncode, 0, out.stdout + out.stderr)


class InvalidInput(unittest.TestCase):
    def test_missing_gate_file_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {})
            rules = json.loads((skill / "references" / "hr-gate-matrix.json").read_text())
            rules["rules"].append({"id": "HR-99", "enforcement": "gate",
                                   "asset": "assets/ghost.py", "test": "assets/ghost.py"})
            (skill / "references" / "hr-gate-matrix.json").write_text(json.dumps(rules))
            code, out = run_gate(skill)
            self.assertEqual(code, 2)

    def test_unreadable_matrix_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            skill = build_skill(Path(d), {"assets/a.py": "x = 1\n"})
            (skill / "references" / "hr-gate-matrix.json").write_text("{broken")
            code, out = run_gate(skill)
            self.assertEqual(code, 2)

    def test_missing_skill_root_exits_two(self):
        code, out = run_gate(Path("/no/such/skill"))
        self.assertEqual(code, 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
