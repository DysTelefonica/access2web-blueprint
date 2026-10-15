#!/usr/bin/env python3
# ci-pattern asset tests — required-suites aggregator (DysTelefonica/team-skills#528)
"""Executable suite for the required-suites gate (stdlib only)."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "check_required_suites.py"
EXAMPLE_POLICY = ASSET / "required-suites.policy.example.json"

POLICY_SUITES = ["lint", "unit", "integration", "e2e-smoke"]


def base_policy() -> dict:
    return {
        "required_suites": POLICY_SUITES,
        "events": {
            "pull_request": {"accepted_unran": ["e2e-smoke"]},
            "push": {"accepted_unran": []},
            "workflow_dispatch": {"accepted_unran": []},
        },
    }


def evidence(items: list[dict] | None = None, measured: bool = True,
             subjects: int | None = None) -> dict:
    items = list(items) if items is not None else []
    return {
        "measured": measured,
        "subjects": subjects if subjects is not None else max(
            sum(e["ran"] + e["skipped"] for e in items), 1),
        "suites": list(items),
    }


def suite(name: str, ran: int = 1, passed: int | None = None,
          skipped: int = 0) -> dict:
    return {"name": name, "ran": ran,
            "passed": passed if passed is not None else ran, "skipped": skipped}


def run_gate(policy: dict | None = None, evidence_doc: dict | None = None,
             event: str = "pull_request",
             policy_bytes: bytes | None = None,
             evidence_bytes: bytes | None = None,
             base_policy_doc: dict | None = None) -> subprocess.CompletedProcess:
    """Run the gate as CI does."""
    with tempfile.TemporaryDirectory() as tmp:
        policy_path = Path(tmp) / "policy.json"
        if policy_bytes is None:
            policy_bytes = json.dumps(
                policy if policy is not None else base_policy()).encode("utf-8")
        policy_path.write_bytes(policy_bytes)
        argv = [sys.executable, str(SCRIPT), "--policy", str(policy_path),
                "--event", event]
        if base_policy_doc is not None:
            base_path = Path(tmp) / "base-policy.json"
            base_path.write_text(json.dumps(base_policy_doc), encoding="utf-8")
            argv += ["--base-policy", str(base_path)]
        evidence_path = Path(tmp) / "evidence.json"
        if evidence_bytes is None:
            evidence_bytes = json.dumps(
                evidence_doc if evidence_doc is not None else evidence()).encode("utf-8")
        evidence_path.write_bytes(evidence_bytes)
        argv += ["--evidence", str(evidence_path)]
        proc = subprocess.run(argv, capture_output=True)
        proc.stdout = proc.stdout.decode("utf-8", errors="replace")
        proc.stderr = proc.stderr.decode("utf-8", errors="replace")
        return proc


class CleanTests(unittest.TestCase):
    def test_all_required_suites_ran_exits_zero(self) -> None:
        items = [suite("lint", ran=4), suite("unit", ran=87),
                 suite("integration", ran=12), suite("e2e-smoke", ran=2)]
        proc = run_gate(evidence_doc=evidence(items))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("VERDICT: PASS", proc.stdout)
        self.assertIn("lint (ran=4, passed=4, skipped=0)", proc.stdout)

    def test_accepted_unran_is_not_a_failure(self) -> None:
        items = [suite("lint", ran=2), suite("unit", ran=20),
                 suite("integration", ran=8)]
        proc = run_gate(evidence_doc=evidence(items))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("e2e-smoke (accepted unran for 'pull_request')", proc.stdout)

    def test_extra_suites_in_evidence_are_ignored(self) -> None:
        items = [suite("lint", ran=2), suite("unit", ran=10),
                 suite("integration", ran=4), suite("e2e-smoke", ran=1),
                 suite("ratchet", ran=1)]
        proc = run_gate(evidence_doc=evidence(items))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


class VerdictTests(unittest.TestCase):
    def test_required_suite_absent_from_evidence_exits_nonzero(self) -> None:
        items = [suite("lint", ran=2), suite("unit", ran=20)]
        proc = run_gate(evidence_doc=evidence(items))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("integration", proc.stdout)
        self.assertIn("e2e-smoke", proc.stdout)

    def test_zero_run_outside_accepted_unran_exits_nonzero(self) -> None:
        items = [suite("lint", ran=2), suite("unit", ran=10),
                 suite("integration", ran=0), suite("e2e-smoke", ran=1)]
        proc = run_gate(evidence_doc=evidence(items), event="push")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("integration", proc.stdout)
        self.assertIn("ran zero tests", proc.stdout)

    def test_zero_run_for_e2e_smoke_on_pull_request_is_accepted(self) -> None:
        items = [suite("lint", ran=2), suite("unit", ran=10),
                 suite("integration", ran=4), suite("e2e-smoke", ran=0)]
        proc = run_gate(evidence_doc=evidence(items))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_failed_run_with_zero_passes_exits_nonzero(self) -> None:
        items = [suite("lint", ran=5, passed=0), suite("unit", ran=10),
                 suite("integration", ran=4), suite("e2e-smoke", ran=1)]
        proc = run_gate(evidence_doc=evidence(items))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("ran tests but passed none", proc.stdout)
        self.assertIn("lint", proc.stdout)

    def test_undeclared_event_exits_nonzero(self) -> None:
        items = [suite("lint", ran=1), suite("unit", ran=1),
                 suite("integration", ran=1), suite("e2e-smoke", ran=1)]
        proc = run_gate(evidence_doc=evidence(items), event="schedule")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("event 'schedule' is not declared", proc.stdout)


class LivenessTests(unittest.TestCase):
    def test_measured_false_exits_two(self) -> None:
        doc = evidence([suite("lint", ran=2)], measured=False)
        proc = run_gate(evidence_doc=doc)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("measured", proc.stderr)

    def test_subjects_zero_exits_two(self) -> None:
        doc = evidence([suite("lint", ran=2)], subjects=0)
        proc = run_gate(evidence_doc=doc)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("subjects", proc.stderr)

    def test_extra_top_level_key_in_evidence_exits_two(self) -> None:
        raw = json.dumps({"measured": True, "subjects": 1,
                          "suites": [], "findings": []}).encode("utf-8")
        proc = run_gate(evidence_bytes=raw)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("must carry exactly", proc.stderr)

    def test_malformed_evidence_json_exits_two(self) -> None:
        proc = run_gate(evidence_bytes=b"{not-json")
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("not valid JSON", proc.stderr)


class PolicyValidationTests(unittest.TestCase):
    def test_duplicate_suite_in_policy_exits_two(self) -> None:
        policy = base_policy()
        policy["required_suites"] = ["lint", "unit", "unit", "integration"]
        proc = run_gate(policy=policy)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("duplicate", proc.stderr)

    def test_outsider_in_accepted_unran_exits_two(self) -> None:
        policy = base_policy()
        policy["events"]["pull_request"]["accepted_unran"] = ["ratchet"]
        proc = run_gate(policy=policy)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("outside required_suites", proc.stderr)

    def test_relaxation_without_identity_exits_two(self) -> None:
        policy = base_policy()
        policy["relaxation"] = {"reason": "x", "reference": "y"}
        proc = run_gate(policy=policy)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("verifiable identity", proc.stderr)

    def test_negative_ran_in_evidence_exits_two(self) -> None:
        doc = evidence([{"name": "lint", "ran": -1, "passed": 0, "skipped": 0}])
        proc = run_gate(evidence_doc=doc)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("ran", proc.stderr)


class BaseComparisonTests(unittest.TestCase):
    def test_suite_removed_without_relaxation_fails(self) -> None:
        items = [suite("lint", ran=2), suite("unit", ran=10),
                 suite("integration", ran=4), suite("e2e-smoke", ran=1)]
        base = {
            "required_suites": ["lint", "unit", "integration", "e2e-smoke",
                                "ratchet"],
            "events": base_policy()["events"],
        }
        proc = run_gate(evidence_doc=evidence(items), base_policy_doc=base)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("REFUSED", proc.stdout)
        self.assertIn("'ratchet' removed from required_suites", proc.stdout)

    def test_suite_removed_with_relaxation_passes(self) -> None:
        items = [suite("lint", ran=2), suite("unit", ran=10),
                 suite("integration", ran=4), suite("e2e-smoke", ran=1)]
        base = {
            "required_suites": ["lint", "unit", "integration", "e2e-smoke",
                                "ratchet"],
            "events": base_policy()["events"],
        }
        candidate = base_policy()
        candidate["relaxation"] = {
            "reason": "ratchet opt-in only triggers on first big refactor",
            "reference": "DysTelefonica/team-skills#528",
            "issue": 528,
        }
        proc = run_gate(policy=candidate, evidence_doc=evidence(items),
                         base_policy_doc=base)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("covered by relaxation", proc.stdout)

    def test_inherited_relaxation_field_is_refused(self) -> None:
        items = [suite("lint", ran=2), suite("unit", ran=10),
                 suite("integration", ran=4), suite("e2e-smoke", ran=1)]
        shared = {
            "reason": "shared reason from base",
            "reference": "DysTelefonica/team-skills#500",
            "issue": 500,
        }
        base = {
            "required_suites": ["lint", "unit", "integration", "e2e-smoke",
                                "ratchet"],
            "events": base_policy()["events"],
            "relaxation": shared,
        }
        candidate = base_policy()
        candidate["relaxation"] = dict(shared)
        proc = run_gate(policy=candidate, evidence_doc=evidence(items),
                         base_policy_doc=base)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("inherited field does not cover", proc.stdout)

    def test_no_base_policy_declares_comparison_not_performed(self) -> None:
        items = [suite("lint", ran=2), suite("unit", ran=10),
                 suite("integration", ran=4), suite("e2e-smoke", ran=1)]
        proc = run_gate(evidence_doc=evidence(items))
        self.assertIn("base comparison: NOT performed", proc.stdout)


class ExamplePolicyTests(unittest.TestCase):
    def test_example_policy_loads(self) -> None:
        doc = json.loads(EXAMPLE_POLICY.read_text(encoding="utf-8"))
        self.assertEqual(set(doc), {"required_suites", "events"})
        self.assertEqual(doc["required_suites"], POLICY_SUITES)
        self.assertEqual(doc["events"]["pull_request"]["accepted_unran"],
                         ["e2e-smoke"])

    def test_example_policy_with_clean_evidence_exits_zero(self) -> None:
        items = [suite("lint", ran=2), suite("unit", ran=10),
                 suite("integration", ran=4), suite("e2e-smoke", ran=2)]
        proc = run_gate(policy=json.loads(EXAMPLE_POLICY.read_text(encoding="utf-8")),
                         evidence_doc=evidence(items))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main()