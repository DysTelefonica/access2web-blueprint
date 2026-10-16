#!/usr/bin/env python3
# ci-pattern asset tests — knip -> ratchet normalizer (DysTelefonica/team-skills#528)
"""Executable suite for the knip normalizer (stdlib only).

Covers the contract: rule mapping per knip ``type``, fingerprint
stability across re-runs (HR-16), liveness trio (HR-3) and the
failure shapes that exit 2.
"""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "normalize_knip.py"


def fingerprint(rule: str, path: str, symbol: str) -> str:
    return hashlib.sha1(f"{rule}|{path}|{symbol}".encode("utf-8")).hexdigest()


def run_adapter(raw_bytes: bytes) -> subprocess.CompletedProcess:
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "knip.json"
        path.write_bytes(raw_bytes)
        proc = subprocess.run(
            [sys.executable, str(SCRIPT), "--knip-output", str(path)],
            capture_output=True)
    proc.stdout = proc.stdout.decode("utf-8", errors="replace")
    proc.stderr = proc.stderr.decode("utf-8", errors="replace")
    return proc


class CleanTests(unittest.TestCase):
    def test_empty_issues_emits_zero_debt_shape(self) -> None:
        proc = run_adapter(b'{"files": [], "issues": []}')
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        doc = json.loads(proc.stdout)
        self.assertEqual(doc, {"measured": True, "subjects": 0, "findings": []})

    def test_known_knip_types_normalize_to_canonical_rules(self) -> None:
        cases = [
            ("files", "dead-code:unused-files"),
            ("deps", "dead-code:unused-deps"),
            ("devDeps", "dead-code:unused-dev-deps"),
            ("exports", "dead-code:unused-exports"),
            ("nsExports", "dead-code:unused-ns-exports"),
            ("classMembers", "dead-code:unused-class-members"),
            ("duplicates", "dead-code:duplicate-export"),
            ("unresolved", "dead-code:unresolved-import"),
            ("binaries", "dead-code:unused-binary"),
        ]
        payload = {"files": [], "issues": [
            {"file": f"src/{name}.ts", "type": name, "symbol": name} for name, _ in cases
        ]}
        proc = run_adapter(json.dumps(payload).encode("utf-8"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        doc = json.loads(proc.stdout)
        rules = {finding["rule"] for finding in doc["findings"]}
        for _, expected in cases:
            self.assertIn(expected, rules)

    def test_symbol_drives_fingerprint(self) -> None:
        payload = {"files": [], "issues": [
            {"file": "src/foo.ts", "type": "exports", "symbol": "alpha"},
            {"file": "src/foo.ts", "type": "exports", "symbol": "beta"},
        ]}
        proc = run_adapter(json.dumps(payload).encode("utf-8"))
        doc = json.loads(proc.stdout)
        fingerprints = {f["fingerprint"] for f in doc["findings"]}
        self.assertEqual(len(fingerprints), 2)

    def test_subjects_counts_distinct_files_across_files_and_issues(self) -> None:
        payload = {"files": ["src/a.ts", "src/b.ts"], "issues": [
            {"file": "src/a.ts", "type": "files"},
            {"file": "src/c.ts", "type": "exports", "symbol": "x"},
        ]}
        proc = run_adapter(json.dumps(payload).encode("utf-8"))
        doc = json.loads(proc.stdout)
        self.assertEqual(doc["subjects"], 3)

    def test_unknown_knip_type_still_normalizes(self) -> None:
        proc = run_adapter(json.dumps({
            "files": [], "issues": [
                {"file": "src/foo.ts", "type": "some-future-type", "symbol": "x"},
            ]}).encode("utf-8"))
        doc = json.loads(proc.stdout)
        self.assertEqual(doc["findings"][0]["rule"],
                         "dead-code:some-future-type")


class StabilityTests(unittest.TestCase):
    def test_fingerprint_stable_across_runs(self) -> None:
        # HR-16: reordering or shifting findings must not change identities.
        before = {"files": [], "issues": [
            {"file": "src/foo.ts", "type": "exports", "symbol": "bar",
             "line": 100, "col": 4},
        ]}
        after = {"files": [], "issues": [
            {"file": "src/foo.ts", "type": "exports", "symbol": "bar",
             "line": 250, "col": 12},
        ]}
        proc_before = run_adapter(json.dumps(before).encode("utf-8"))
        proc_after = run_adapter(json.dumps(after).encode("utf-8"))
        self.assertEqual(
            json.loads(proc_before.stdout)["findings"][0]["fingerprint"],
            json.loads(proc_after.stdout)["findings"][0]["fingerprint"],
        )


class FailClosedTests(unittest.TestCase):
    def test_non_object_input_exits_two(self) -> None:
        proc = run_adapter(b'[1, 2, 3]')
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("JSON object", proc.stderr)

    def test_missing_files_array_exits_two(self) -> None:
        proc = run_adapter(b'{"issues": []}')
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("'files'", proc.stderr)

    def test_missing_issues_array_exits_two(self) -> None:
        proc = run_adapter(b'{"files": []}')
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("'issues'", proc.stderr)

    def test_non_string_file_exits_two(self) -> None:
        proc = run_adapter(b'{"files": [42], "issues": []}')
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("non-empty strings", proc.stderr)

    def test_missing_file_field_in_issue_exits_two(self) -> None:
        proc = run_adapter(b'{"files": [], "issues": [{"type": "files"}]}')
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("'file'", proc.stderr)

    def test_malformed_json_exits_two(self) -> None:
        proc = run_adapter(b"{not-json")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("not valid JSON", proc.stderr)


class IntegrationTests(unittest.TestCase):
    """The output must be consumable by the ratchet gate itself."""

    def test_normalized_output_passes_clean_ratchet(self) -> None:
        ratchet = ASSET / "check_ratchet.py"
        if not ratchet.exists():
            self.skipTest("check_ratchet.py not present in this branch")
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            payload = {"files": ["src/foo.ts"], "issues": [
                {"file": "src/foo.ts", "type": "exports", "symbol": "bar"},
            ]}
            (tmp / "knip.json").write_text(json.dumps(payload), encoding="utf-8")
            findings_path = tmp / "findings.json"
            proc = subprocess.run(
                [sys.executable, str(SCRIPT),
                 "--knip-output", str(tmp / "knip.json"),
                 "--output", str(findings_path)],
                capture_output=True)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            baseline_path = tmp / "baseline.json"
            baseline_path.write_text(json.dumps({
                "baseline_version": 1,
                "entries": [],
            }), encoding="utf-8")
            ratchet_proc = subprocess.run(
                [sys.executable, str(ratchet),
                 "--findings", str(findings_path),
                 "--baseline", str(baseline_path)],
                capture_output=True)
            ratchet_stdout = ratchet_proc.stdout.decode("utf-8", errors="replace")
            ratchet_stderr = ratchet_proc.stderr.decode("utf-8", errors="replace")
            self.assertEqual(ratchet_proc.returncode, 1,
                             ratchet_stdout + ratchet_stderr)
            self.assertIn("HR-15 new finding identity", ratchet_stdout)


if __name__ == "__main__":
    unittest.main()