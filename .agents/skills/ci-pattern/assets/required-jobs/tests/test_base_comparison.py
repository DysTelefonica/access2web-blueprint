#!/usr/bin/env python3
# ci-pattern asset tests — policy compared against the base-branch copy (DysTelefonica/team-skills#161)
"""Executable suite for ``--base-policy`` and the ``relaxation`` field (HR-15).

The run is judged with the candidate policy; the base copy only decides
whether the candidate relaxed anything, and a relaxation passes only with a
``relaxation`` field the base does not already carry. Every refusal asserts a
NON-ZERO exit of the gate run as a subprocess. Stdlib only.
"""

from __future__ import annotations

import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "check_required_jobs.py"
EXAMPLE_POLICY = ASSET / "required-jobs.policy.example.json"

MISSING = object()  # sentinel: pass a --base-policy path that does not exist

RELAXATION = {"reason": "full moved to the nightly run", "reference": "example/repo#12",
              "actor": "maintainer"}


def base_policy() -> dict:
    return {
        "required_jobs": ["lint", "compile", "smoke", "unit", "full"],
        "events": {"pull_request": {"accepted_skips": ["full"]}, "push": {"accepted_skips": []}},
    }


def run_gate(candidate: dict, base: object = None, event: str = "push",
             needs: dict | None = None) -> subprocess.CompletedProcess:
    """Run the gate with every candidate job green unless ``needs`` says so.

    ``base`` is None (no ``--base-policy``), a policy dict, raw bytes, or
    ``MISSING``. With all jobs green, a non-zero exit can only come from the
    policy itself.
    """
    if needs is None:
        needs = {job: {"result": "success", "outputs": {}} for job in candidate["required_jobs"]}
    with tempfile.TemporaryDirectory() as tmp:
        candidate_path = Path(tmp) / "policy.json"
        candidate_path.write_text(json.dumps(candidate), encoding="utf-8")
        argv = [sys.executable, str(SCRIPT), "--policy", str(candidate_path), "--event", event]
        if base is not None:
            base_path = Path(tmp) / "base-policy.json"
            if isinstance(base, bytes):
                base_path.write_bytes(base)
            elif base is not MISSING:
                base_path.write_text(json.dumps(base), encoding="utf-8")
            argv += ["--base-policy", str(base_path)]
        return subprocess.run(argv, input=json.dumps(needs), capture_output=True, text=True)


def without_job(policy: dict, job: str) -> dict:
    relaxed = copy.deepcopy(policy)
    relaxed["required_jobs"].remove(job)
    for spec in relaxed["events"].values():
        if job in spec["accepted_skips"]:
            spec["accepted_skips"].remove(job)
    return relaxed


class DeclaredComparisonTests(unittest.TestCase):
    """The verdict always says whether the base comparison happened (HR-3)."""

    def test_without_base_policy_the_output_declares_it(self) -> None:
        proc = run_gate(base_policy())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("base comparison: NOT performed", proc.stdout)

    def test_with_base_policy_the_output_declares_it(self) -> None:
        proc = run_gate(base_policy(), base_policy())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("base comparison: performed", proc.stdout)
        self.assertIn("no relaxation detected", proc.stdout)
        self.assertNotIn("NOT performed", proc.stdout)

    def test_failing_run_still_declares_the_comparison(self) -> None:
        needs = {job: "success" for job in base_policy()["required_jobs"]}
        needs["unit"] = "failure"
        proc = run_gate(base_policy(), needs=needs)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("base comparison: NOT performed", proc.stdout)

    def test_example_policy_compared_with_itself_passes(self) -> None:
        example = json.loads(EXAMPLE_POLICY.read_text(encoding="utf-8"))
        proc = run_gate(example, example)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


class RelaxationTests(unittest.TestCase):
    """A relaxation passes only with a relaxation field the base lacks."""

    def relaxations(self) -> dict:
        skip_added = base_policy()
        skip_added["events"]["push"]["accepted_skips"] = ["full"]
        event_added = base_policy()
        event_added["events"]["schedule"] = {"accepted_skips": []}
        return {
            "job 'full' removed from required_jobs": without_job(base_policy(), "full"),
            "event 'push': accepted skip 'full' added": skip_added,
            "event 'schedule' added": event_added,
        }

    def test_relaxation_without_field_exits_nonzero(self) -> None:
        for finding, candidate in self.relaxations().items():
            with self.subTest(finding=finding):
                proc = run_gate(candidate, base_policy())
                self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
                self.assertIn("VERDICT: FAIL", proc.stdout)
                self.assertIn(finding, proc.stdout)
                self.assertIn("REFUSED", proc.stdout)

    def test_relaxation_with_new_field_exits_zero(self) -> None:
        for finding, candidate in self.relaxations().items():
            with self.subTest(finding=finding):
                candidate["relaxation"] = RELAXATION
                proc = run_gate(candidate, base_policy())
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
                self.assertIn(finding, proc.stdout)
                self.assertIn(RELAXATION["reason"], proc.stdout)
                self.assertIn(RELAXATION["reference"], proc.stdout)
                self.assertNotIn("REFUSED", proc.stdout)

    def test_every_relaxation_is_listed(self) -> None:
        candidate = without_job(base_policy(), "unit")
        candidate["events"]["push"]["accepted_skips"] = ["full"]
        candidate["events"]["schedule"] = {"accepted_skips": []}
        proc = run_gate(candidate, base_policy())
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("3 relaxations detected", proc.stdout)
        for finding in ("job 'unit' removed from required_jobs",
                        "event 'push': accepted skip 'full' added",
                        "event 'schedule' added"):
            self.assertIn(finding, proc.stdout)

    def test_field_inherited_from_the_base_does_not_cover_a_new_relaxation(self) -> None:
        base = base_policy()
        base["relaxation"] = RELAXATION
        candidate = without_job(base, "full")
        proc = run_gate(candidate, base)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("job 'full' removed from required_jobs", proc.stdout)
        self.assertIn("REFUSED", proc.stdout)
        self.assertIn("base policy already carries", proc.stdout)

    def test_inherited_field_padded_with_whitespace_is_still_inherited(self) -> None:
        base = base_policy()
        base["relaxation"] = RELAXATION
        candidate = without_job(base, "full")
        candidate["relaxation"] = {key: f"  {value}\n" for key, value in RELAXATION.items()}
        proc = run_gate(candidate, base)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("REFUSED", proc.stdout)

    def test_replaced_field_covers_a_new_relaxation(self) -> None:
        base = base_policy()
        base["relaxation"] = RELAXATION
        candidate = without_job(base, "full")
        candidate["relaxation"] = {"reason": "full retired", "reference": "example/repo#40",
                                   "actor": "maintainer"}
        proc = run_gate(candidate, base)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_covered_relaxation_does_not_hide_a_failed_job(self) -> None:
        candidate = without_job(base_policy(), "full")
        candidate["relaxation"] = RELAXATION
        needs = {job: "success" for job in candidate["required_jobs"]}
        needs["lint"] = "failure"
        proc = run_gate(candidate, base_policy(), needs=needs)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("lint (failure)", proc.stdout)


class TighteningTests(unittest.TestCase):
    """Tightening the policy needs no relaxation field."""

    def test_tightening_without_field_exits_zero(self) -> None:
        job_added = base_policy()
        job_added["required_jobs"].append("e2e")
        skip_removed = base_policy()
        skip_removed["events"]["pull_request"]["accepted_skips"] = []
        event_removed = base_policy()
        del event_removed["events"]["pull_request"]
        for name, candidate in (("job added", job_added), ("skip removed", skip_removed),
                                ("event removed", event_removed)):
            with self.subTest(change=name):
                proc = run_gate(candidate, base_policy())
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
                self.assertIn("no relaxation detected", proc.stdout)

    def test_inherited_field_is_accepted_when_nothing_is_relaxed(self) -> None:
        base = base_policy()
        base["relaxation"] = RELAXATION
        candidate = copy.deepcopy(base)
        candidate["required_jobs"].append("e2e")
        proc = run_gate(candidate, base)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


class RelaxationFieldShapeTests(unittest.TestCase):
    """The field is validated whenever present, with or without a base."""

    def assert_fail_closed(self, proc: subprocess.CompletedProcess, fragment: str) -> None:
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("FAIL fail-closed:", proc.stderr)
        self.assertIn(fragment, proc.stderr)
        self.assertNotIn("Traceback", proc.stderr)

    def test_well_formed_field_is_accepted_without_a_base(self) -> None:
        candidate = base_policy()
        candidate["relaxation"] = RELAXATION
        proc = run_gate(candidate)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_malformed_field_exits_nonzero(self) -> None:
        malformed = {
            "not an object": "example/repo#12",
            "missing reference": {"reason": "x"},
            "missing reason": {"reference": "x"},
            "empty reason": {"reason": "", "reference": "x"},
            "blank reference": {"reason": "x", "reference": "  \n"},
            "non-string reason": {"reason": 12, "reference": "x"},
            "extra key": {"reason": "x", "reference": "y", "approved_by": "z"},
            "null": None,
        }
        for name, field in malformed.items():
            for base in (None, base_policy()):
                with self.subTest(field=name, with_base=base is not None):
                    candidate = base_policy()
                    candidate["relaxation"] = field
                    self.assert_fail_closed(run_gate(candidate, base), "relaxation")

    def test_unknown_top_level_key_still_exits_nonzero(self) -> None:
        candidate = base_policy()
        candidate["exemptions"] = ["unit"]
        self.assert_fail_closed(run_gate(candidate), "policy file")


class BasePolicyInputTests(unittest.TestCase):
    """An unusable base copy is a doubt, never a silent skip of the check."""

    def assert_fail_closed(self, proc: subprocess.CompletedProcess, fragment: str) -> None:
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("FAIL fail-closed:", proc.stderr)
        self.assertIn(fragment, proc.stderr)
        self.assertNotIn("Traceback", proc.stderr)
        self.assertNotIn("VERDICT: PASS", proc.stdout)

    def test_missing_base_policy_file_exits_nonzero(self) -> None:
        self.assert_fail_closed(run_gate(base_policy(), MISSING), "unreadable base policy file")

    def test_base_policy_that_is_the_candidate_file_exits_nonzero(self) -> None:
        # Comparing a file with itself finds nothing by construction.
        proc = subprocess.run(
            [sys.executable, str(SCRIPT), "--policy", str(EXAMPLE_POLICY),
             "--base-policy", str(EXAMPLE_POLICY), "--event", "push"],
            input=json.dumps({job: "success" for job in base_policy()["required_jobs"]}),
            capture_output=True, text=True,
        )
        self.assert_fail_closed(proc, "same file")

    def test_empty_base_policy_file_exits_nonzero(self) -> None:
        self.assert_fail_closed(run_gate(base_policy(), b""), "base policy file")

    def test_base_policy_that_is_not_json_exits_nonzero(self) -> None:
        self.assert_fail_closed(run_gate(base_policy(), b"{not json"), "base policy file")

    def test_base_policy_that_is_not_utf8_exits_nonzero(self) -> None:
        self.assert_fail_closed(run_gate(base_policy(), b"\xff\xfe{}"), "base policy file")

    def test_structurally_invalid_base_policy_exits_nonzero(self) -> None:
        invalid = base_policy()
        invalid["events"]["push"] = {}
        self.assert_fail_closed(run_gate(base_policy(), invalid), "base policy file")

    def test_base_policy_with_malformed_relaxation_exits_nonzero(self) -> None:
        invalid = base_policy()
        invalid["relaxation"] = {"reason": "x"}
        self.assert_fail_closed(run_gate(base_policy(), invalid), "base policy file")


if __name__ == "__main__":
    unittest.main(verbosity=2)
