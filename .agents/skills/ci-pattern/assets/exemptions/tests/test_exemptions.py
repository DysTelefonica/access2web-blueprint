#!/usr/bin/env python3
# ci-pattern asset tests — exemption identity gate (HR-33, DysTelefonica/team-skills#245)
"""Executable suite for the exemption identity gate (black-box, like CI).

HR-33: every gate exemption declared as data — a ``relaxation`` object or an
``exemption`` field — MUST carry a verifiable identity: a non-empty
``actor`` string or a positive ``issue`` integer. The gate scans the JSON
documents it is given, recursively, and reports every exemption without
one. Malformed documents exit 2; a document without exemptions is clean
(0); an empty subject (no files) exits 2. The suite also dogfoods the real
matrix: its HR-19/HR-20 dogfooding exemptions carry ``issue`` identities."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_exemptions.py"
MATRIX = ASSET.parents[1] / "references" / "hr-gate-matrix.json"


def run_gate(files, extra=()):
    paths = []
    with tempfile.TemporaryDirectory() as d:
        for i, doc in enumerate(files):
            p = Path(d, f"doc{i}.json")
            if isinstance(doc, str):
                p.write_text(doc, encoding="utf-8")
            elif isinstance(doc, Path):
                paths.append(str(doc))
                continue
            else:
                p.write_text(json.dumps(doc), encoding="utf-8")
            paths.append(str(p))
        out = subprocess.run([sys.executable, str(GATE), *sum(
            (["--file", p] for p in paths), []), *extra],
            capture_output=True, text=True)
    return out.returncode, out.stdout + out.stderr


class Clean(unittest.TestCase):
    def test_relaxation_with_actor_is_clean(self):
        code, out = run_gate([{"relaxation": {"reason": "r", "reference": "x",
                                              "actor": "auditor"}}])
        self.assertEqual((code, "HR-33" in out), (0, False))

    def test_exemption_with_issue_is_clean(self):
        code, out = run_gate([{"exemption": {"reason": "no releases here",
                                             "issue": 195}}])
        self.assertEqual((code, "HR-33" in out), (0, False))

    def test_document_without_exemptions_is_clean(self):
        code, out = run_gate([{"entries": [], "baseline_version": 1}])
        self.assertEqual((code, "HR-33" in out), (0, False))

    def test_nested_exemption_is_found_and_clean(self):
        code, out = run_gate([{"rules": [{"exemption": {"reason": "r",
                                                        "actor": "maintainer"}}]}])
        self.assertEqual((code, "HR-33" in out), (0, False))


class Findings(unittest.TestCase):
    def test_relaxation_without_identity(self):
        code, out = run_gate([{"relaxation": {"reason": "r", "reference": "x"}}])
        self.assertEqual(code, 1)
        self.assertIn("HR-33", out)

    def test_exemption_as_bare_string(self):
        code, out = run_gate([{"exemption": "audited by hand"}])
        self.assertEqual(code, 1)
        self.assertIn("HR-33", out)

    def test_exemption_object_without_identity(self):
        code, out = run_gate([{"exemption": {"reason": "r"}}])
        self.assertEqual(code, 1)
        self.assertIn("HR-33", out)

    def test_exemption_with_empty_actor(self):
        code, out = run_gate([{"exemption": {"reason": "r", "actor": ""}}])
        self.assertEqual(code, 1)

    def test_exemption_with_non_positive_issue(self):
        code, out = run_gate([{"exemption": {"reason": "r", "issue": 0}}])
        self.assertEqual(code, 1)

    def test_findings_aggregate_across_files(self):
        code, out = run_gate([
            {"relaxation": {"reason": "r", "reference": "x"}},
            {"exemption": {"reason": "r", "actor": ""}},
        ])
        self.assertEqual(code, 1)
        self.assertIn("doc0.json", out)
        self.assertIn("doc1.json", out)

    def test_degenerate_deep_nesting_still_found(self):
        code, out = run_gate([{"a": {"b": {"c": [{"exemption": {"reason": "r"}}]}}}])
        self.assertEqual(code, 1)
        self.assertIn("HR-33", out)


class InvalidInput(unittest.TestCase):
    def test_no_files_exits_two(self):
        code, out = run_gate([])
        self.assertEqual(code, 2)

    def test_unreadable_file_exits_two(self):
        code, out = run_gate(["not json at all"])
        self.assertEqual(code, 2)

    def test_missing_file_exits_two(self):
        code, out = run_gate([], extra=("--file", "/no/such.json"))
        self.assertEqual(code, 2)

    def test_non_object_document_exits_two(self):
        code, out = run_gate(["[1, 2]"])
        self.assertEqual(code, 2)


class Dogfooding(unittest.TestCase):
    def test_real_documents_exemptions_carry_identity(self):
        """Dogfooding: la matriz HR→gate y los JSON de .github/ de este repo
        pasan el escaneo de identidad de exenciones."""
        repo_root = ASSET.parents[4]
        docs = [MATRIX,
                repo_root / ".github" / "host-contract.json",
                repo_root / ".github" / "workflow-policy.json"]
        code, out = run_gate(docs)
        self.assertEqual(code, 0, out)


if __name__ == "__main__":
    unittest.main(verbosity=2)
