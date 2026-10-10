#!/usr/bin/env python3
# ci-pattern asset tests — runbook premises gate (HR-42, DysTelefonica/team-skills#167)
"""Executable suite for the runbook premises gate (black-box).

HR-42: every state premise a runbook assumes MUST be seeded idempotently
or verified fail-loud with the verification command written in the runbook
itself — prose like «is confirmed seeded» without a command behind it is a
runbook defect, not production state. The gate scans runbook markdown:
premise lines (matched against versioned patterns, case-insensitive) must
be backed by at least one line in the same file referencing a resource
token of the premise (single-quoted strings on the premise line).

Exit codes: 0 clean · 1 findings · 2 invalid input (missing root, missing
or malformed patterns, empty subject — a root with no runbook files)."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_runbook_premises.py"
PATTERNS = {"premise_patterns": ["confirmed seeded", "se asume sembrado"]}


def build(root: Path, files: dict):
    for rel, content in files.items():
        p = root / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")


def run_gate(root, patterns=PATTERNS, extra=()):
    with tempfile.TemporaryDirectory() as ad:
        pp = Path(ad, "patterns.json")
        pp.write_text(json.dumps(patterns), encoding="utf-8")
        out = subprocess.run([sys.executable, str(GATE), "--root", str(root),
                              "--patterns", str(pp), *extra],
                             capture_output=True, text=True)
    return out.returncode, out.stdout + out.stderr


class Clean(unittest.TestCase):
    def test_premise_with_verification_in_runbook_is_clean(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            build(root, {"RELEASE-v1.0.0.md":
                         "# Runbook\n\nThe e2e user 'e2e@apap.local' is confirmed seeded.\n\n"
                         "```bash\npsql -c \"SELECT 1 FROM users WHERE email = 'e2e@apap.local'\"\n"
                         "```\n"})
            code, out = run_gate(root)
            self.assertEqual((code, "HR-42" in out), (0, False))

    def test_premise_without_patterns_is_ignored(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            build(root, {"r.md": "The user is confirmed seeded.\n"})
            code, out = run_gate(root, {"premise_patterns": ["otra cosa"]})
            self.assertEqual((code, "HR-42" in out), (0, False))

    def test_case_insensitive_matching(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            build(root, {"r.md": "Is Confirmed Seeded: 'e2e@apap.local'\n\n"
                                 "```bash\necho 'e2e@apap.local'\n```\n"})
            code, out = run_gate(root)
            self.assertEqual((code, "HR-42" in out), (0, False))


class Findings(unittest.TestCase):
    def test_premise_without_verification_is_a_finding(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            build(root, {"RELEASE-v1.0.0.md":
                         "The e2e user is confirmed seeded.\n\n"
                         "```bash\n./deploy.sh\n```\n"})
            code, out = run_gate(root)
            self.assertEqual(code, 1)
            self.assertIn("HR-42", out)

    def test_premise_without_resource_token_is_a_finding(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            build(root, {"r.md": "The e2e user is confirmed seeded.\n\n"
                                 "```bash\necho hi\n```\n"})
            code, out = run_gate(root)
            self.assertEqual(code, 1)
            self.assertIn("resource", out)

    def test_comment_only_code_block_is_not_a_verification(self):
        """Hallazgo del auditor #268: un bloque cuyo contenido es solo
        comentario no ejecuta nada — la premisa sigue sin comando."""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            build(root, {"RELEASE-v1.0.0.md":
                         "The e2e user 'e2e@apap.local' is confirmed seeded.\n\n"
                         "```bash\n# 'e2e@apap.local' se asume presente\n```\n"})
            code, out = run_gate(root)
            self.assertEqual(code, 1)
            self.assertIn("HR-42", out)

    def test_verification_in_other_file_does_not_count(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            build(root, {"r1.md": "confirmed seeded: 'e2e@apap.local'\n",
                         "r2.md": "```bash\necho 'e2e@apap.local'\n```\n"})
            code, out = run_gate(root)
            self.assertEqual(code, 1)
            self.assertIn("r1.md", out)

    def test_findings_cite_file_and_line(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            build(root, {"docs/deep/r.md": "line 3 premise 'e2e@apap.local' is confirmed seeded\n"})
            code, out = run_gate(root)
            self.assertEqual(code, 1)
            self.assertIn("docs/deep/r.md:1", out)


class InvalidInput(unittest.TestCase):
    def test_missing_root_exits_two(self):
        code, out = run_gate(Path("/no/such/root"))
        self.assertEqual(code, 2)

    def test_malformed_patterns_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            build(root, {"r.md": "hi\n"})
            pp = Path(root, "patterns.json")
            pp.write_text("{broken", encoding="utf-8")
            out = subprocess.run([sys.executable, str(GATE), "--root", str(root),
                                  "--patterns", str(pp)], capture_output=True, text=True)
            self.assertEqual(out.returncode, 2)

    def test_patterns_without_list_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            build(root, {"r.md": "hi\n"})
            code, out = run_gate(root, {"otra_clave": 1})
            self.assertEqual(code, 2)

    def test_empty_subject_exits_two(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            code, out = run_gate(root)
            self.assertEqual(code, 2)

    def test_missing_required_argument(self):
        out = subprocess.run([sys.executable, str(GATE)], capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
