#!/usr/bin/env python3
# ci-pattern asset tests — deploy review policy gate (HR-40/HR-41, DysTelefonica/team-skills#166)
"""Executable suite for the deploy review policy gate (black-box).

HR-40: the policy MUST declare the resolution source of the previous
review; a missing or unknown source is a finding that names the cause —
never a silent fallback. HR-41: bootstrap exemptions MUST be symmetric
across the declared required contexts — one context exempt and another
without declared coverage is a violation. Fail-closed: unreadable
documents, malformed structures or an empty context map exit 2.

Exit codes: 0 clean · 1 findings · 2 invalid input or empty subject."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_deploy_policy.py"

POLICY = {
    "policy_version": 1,
    "resolution_source": {"kind": "health-endpoint", "location": "https://host/healthz"},
    "contexts": {
        "release/e2e-production": {"required": True, "bootstrap_exempt": True},
        "release/smoke-production": {"required": True, "bootstrap_exempt": True},
    },
}


def run_gate(policy, extra=()):
    with tempfile.TemporaryDirectory() as d:
        p = Path(d, "policy.json")
        if isinstance(policy, str):
            p.write_text(policy, encoding="utf-8")
        else:
            p.write_text(json.dumps(policy), encoding="utf-8")
        out = subprocess.run([sys.executable, str(GATE), "--policy", str(p), *extra],
                             capture_output=True, text=True)
    return out.returncode, out.stdout + out.stderr


class Clean(unittest.TestCase):
    def test_symmetric_bootstrap_exemptions_are_clean(self):
        code, out = run_gate(POLICY)
        self.assertEqual((code, "HR-40" in out), (0, False))

    def test_symmetric_without_bootstrap_is_clean(self):
        policy = json.loads(json.dumps(POLICY))
        for c in policy["contexts"].values():
            c["bootstrap_exempt"] = False
        code, out = run_gate(policy)
        self.assertEqual((code, "HR-40" in out), (0, False))


class HR40ResolutionSource(unittest.TestCase):
    def test_missing_resolution_source_fails_naming_the_cause(self):
        policy = json.loads(json.dumps(POLICY))
        del policy["resolution_source"]
        code, out = run_gate(policy)
        self.assertEqual(code, 1)
        self.assertIn("HR-40", out)
        self.assertIn("resolution_source", out)

    def test_unknown_resolution_kind_fails_naming_the_cause(self):
        policy = json.loads(json.dumps(POLICY))
        policy["resolution_source"]["kind"] = "run-list-heuristic"
        code, out = run_gate(policy)
        self.assertEqual(code, 1)
        self.assertIn("run-list-heuristic", out)

    def test_empty_location_fails_naming_the_cause(self):
        policy = json.loads(json.dumps(POLICY))
        policy["resolution_source"]["location"] = ""
        code, out = run_gate(policy)
        self.assertEqual(code, 1)

    def test_declared_source_is_clean(self):
        code, out = run_gate(POLICY)
        self.assertEqual((code, "HR-40" in out), (0, False))


class HR41BootstrapSymmetry(unittest.TestCase):
    def test_asymmetric_exemptions_fail(self):
        policy = json.loads(json.dumps(POLICY))
        policy["contexts"]["release/smoke-production"]["bootstrap_exempt"] = False
        code, out = run_gate(policy)
        self.assertEqual(code, 1)
        self.assertIn("HR-41", out)

    def test_context_without_declared_coverage_fails(self):
        policy = json.loads(json.dumps(POLICY))
        del policy["contexts"]["release/smoke-production"]["bootstrap_exempt"]
        code, out = run_gate(policy)
        self.assertEqual(code, 1)
        self.assertIn("HR-41", out)

    def test_negative_symmetric_coverage_stays_clean(self):
        policy = json.loads(json.dumps(POLICY))
        policy["contexts"]["release/e2e-production"]["bootstrap_exempt"] = False
        policy["contexts"]["release/smoke-production"]["bootstrap_exempt"] = False
        code, out = run_gate(policy)
        self.assertEqual(code, 0)


class InvalidInput(unittest.TestCase):
    def test_missing_policy_exits_two(self):
        code, out = run_gate("not json at all")
        self.assertEqual(code, 2)

    def test_non_object_policy_exits_two(self):
        code, out = run_gate("[1, 2]")
        self.assertEqual(code, 2)

    def test_unknown_policy_key_exits_two(self):
        policy = json.loads(json.dumps(POLICY))
        policy["extra"] = 1
        code, out = run_gate(policy)
        self.assertEqual(code, 2)

    def test_empty_contexts_exits_two(self):
        policy = json.loads(json.dumps(POLICY))
        policy["contexts"] = {}
        code, out = run_gate(policy)
        self.assertEqual(code, 2)

    def test_context_entry_not_an_object_exits_two(self):
        policy = json.loads(json.dumps(POLICY))
        policy["contexts"]["release/e2e-production"] = "required"
        code, out = run_gate(policy)
        self.assertEqual(code, 2)

    def test_context_missing_required_flag_exits_two(self):
        policy = json.loads(json.dumps(POLICY))
        del policy["contexts"]["release/e2e-production"]["required"]
        code, out = run_gate(policy)
        self.assertEqual(code, 2)

    def test_missing_required_argument_exits_two(self):
        out = subprocess.run([sys.executable, str(GATE)], capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
