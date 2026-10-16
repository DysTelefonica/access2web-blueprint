#!/usr/bin/env python3
# ci-pattern asset tests — vulture -> ratchet normalizer (DysTelefonica/team-skills#528)
"""Executable suite for the vulture normalizer (stdlib only).

Covers the contract of the normalized output: rule mapping, fingerprint
stability across re-runs (HR-16), liveness trio (HR-3) and the failure
shapes that exit 2.
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
SCRIPT = ASSET / "normalize_vulture.py"


def fingerprint(rule: str, path: str, symbol: str) -> str:
    return hashlib.sha1(f"{rule}|{path}|{symbol}".encode("utf-8")).hexdigest()


def run_adapter(raw_bytes: bytes) -> subprocess.CompletedProcess:
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "vulture.json"
        path.write_bytes(raw_bytes)
        proc = subprocess.run(
            [sys.executable, str(SCRIPT), "--vulture-output", str(path)],
            capture_output=True)
    proc.stdout = proc.stdout.decode("utf-8", errors="replace")
    proc.stderr = proc.stderr.decode("utf-8", errors="replace")
    return proc


class CleanTests(unittest.TestCase):
    def test_empty_array_emits_zero_debt_shape(self) -> None:
        proc = run_adapter(b"[]")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        doc = json.loads(proc.stdout)
        self.assertEqual(doc, {"measured": True, "subjects": 0, "findings": []})

    def test_unused_function_normalizes_with_correct_rule(self) -> None:
        payload = json.dumps([{
            "file": "src/foo.py", "type": "unused-function",
            "message": "unused function 'bar'", "first_line": 42,
            "function": "bar",
        }]).encode("utf-8")
        proc = run_adapter(payload)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        doc = json.loads(proc.stdout)
        self.assertEqual(doc["measured"], True)
        self.assertEqual(doc["subjects"], 1)
        self.assertEqual(doc["findings"], [{
            "rule": "dead-code:unused-function",
            "path": "src/foo.py",
            "fingerprint": fingerprint("dead-code:unused-function",
                                       "src/foo.py", "bar"),
        }])

    def test_multiple_findings_keep_distinct_fingerprints(self) -> None:
        payload = json.dumps([
            {"file": "src/foo.py", "type": "unused-function", "function": "bar"},
            {"file": "src/foo.py", "type": "unused-function", "function": "baz"},
            {"file": "src/qux.py", "type": "unused-function", "function": "bar"},
        ]).encode("utf-8")
        proc = run_adapter(payload)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        doc = json.loads(proc.stdout)
        self.assertEqual(doc["subjects"], 2)  # two distinct files
        fingerprints = {f["fingerprint"] for f in doc["findings"]}
        self.assertEqual(len(fingerprints), 3)

    def test_unknown_vulture_type_still_normalizes(self) -> None:
        payload = json.dumps([{
            "file": "src/foo.py", "type": "some-future-type", "function": "x",
        }]).encode("utf-8")
        proc = run_adapter(payload)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        doc = json.loads(proc.stdout)
        self.assertEqual(doc["findings"][0]["rule"], "dead-code:some-future-type")

    def test_unnamed_finding_uses_empty_symbol_in_fingerprint(self) -> None:
        payload = json.dumps([{
            "file": "src/foo.py", "type": "unused-import",
        }]).encode("utf-8")
        proc = run_adapter(payload)
        doc = json.loads(proc.stdout)
        self.assertEqual(doc["findings"][0]["fingerprint"],
                         fingerprint("dead-code:unused-import", "src/foo.py", ""))


class StabilityTests(unittest.TestCase):
    def test_fingerprint_does_not_change_with_line_number(self) -> None:
        # HR-16: reordering or shifting findings must not change identities.
        before = json.dumps([{
            "file": "src/foo.py", "type": "unused-function",
            "first_line": 100, "function": "bar",
        }]).encode("utf-8")
        after = json.dumps([{
            "file": "src/foo.py", "type": "unused-function",
            "first_line": 250, "function": "bar",
        }]).encode("utf-8")
        proc_before = run_adapter(before)
        proc_after = run_adapter(after)
        doc_before = json.loads(proc_before.stdout)
        doc_after = json.loads(proc_after.stdout)
        self.assertEqual(doc_before["findings"][0]["fingerprint"],
                         doc_after["findings"][0]["fingerprint"])


class FailClosedTests(unittest.TestCase):
    def test_non_array_input_exits_two(self) -> None:
        proc = run_adapter(b'{"file": "src/foo.py"}')
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("JSON array", proc.stderr)

    def test_missing_file_field_exits_two(self) -> None:
        proc = run_adapter(b'[{"type": "unused-function"}]')
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("'file'", proc.stderr)

    def test_missing_type_field_exits_two(self) -> None:
        proc = run_adapter(b'[{"file": "src/foo.py"}]')
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("'type'", proc.stderr)

    def test_non_object_finding_exits_two(self) -> None:
        proc = run_adapter(b'["src/foo.py"]')
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("not an object", proc.stderr)

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
            vulture_payload = json.dumps([{
                "file": "src/foo.py", "type": "unused-function", "function": "bar",
            }]).encode("utf-8")
            (tmp / "vulture.json").write_bytes(vulture_payload)
            findings_path = tmp / "findings.json"
            proc = subprocess.run(
                [sys.executable, str(SCRIPT),
                 "--vulture-output", str(tmp / "vulture.json"),
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
            self.assertEqual(ratchet_proc.returncode, 1,  # finding not in baseline
                             ratchet_stdout + ratchet_stderr)
            self.assertIn("HR-15 new finding identity", ratchet_stdout)


if __name__ == "__main__":
    unittest.main()