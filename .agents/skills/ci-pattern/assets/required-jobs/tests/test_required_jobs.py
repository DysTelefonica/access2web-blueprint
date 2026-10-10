#!/usr/bin/env python3
# ci-pattern asset tests — verdict, needs shapes and fail-closed validation (DysTelefonica/team-skills#132, #161)
"""Executable suite for the required-jobs aggregator.

Every failure case asserts a NON-ZERO exit; the pass cases assert exit 0.
Runs with the stdlib only, the same way CI invokes it:
``python3 personal/ardelperal/ci-pattern/assets/required-jobs/tests/test_required_jobs.py``.
"""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "check_required_jobs.py"
EXAMPLE_POLICY = ASSET / "required-jobs.policy.example.json"

POLICY_JOBS = ["lint", "compile", "smoke", "unit", "full"]


def run_gate(policy: dict, needs: dict | None = None, event: str = "pull_request",
             stdin_text: str | None = None, needs_file: str | None = None,
             policy_bytes: bytes | None = None,
             stdin_bytes: bytes | None = None) -> subprocess.CompletedProcess:
    """Run the gate as CI does. ``policy_bytes`` and ``stdin_bytes`` feed raw
    bytes, for the inputs that no JSON or text value can express."""
    with tempfile.TemporaryDirectory() as tmp:
        policy_path = Path(tmp) / "policy.json"
        if policy_bytes is None:
            policy_bytes = json.dumps(policy).encode("utf-8")
        policy_path.write_bytes(policy_bytes)
        argv = [sys.executable, str(SCRIPT), "--policy", str(policy_path), "--event", event]
        if needs_file is not None:
            argv += ["--needs-file", needs_file]
        elif stdin_bytes is None:
            if stdin_text is None and needs is not None:
                stdin_text = json.dumps(needs)
            if stdin_text is not None:
                stdin_bytes = stdin_text.encode("utf-8")
        proc = subprocess.run(argv, input=stdin_bytes, capture_output=True)
        proc.stdout = proc.stdout.decode("utf-8", errors="replace")
        proc.stderr = proc.stderr.decode("utf-8", errors="replace")
        return proc


def base_policy() -> dict:
    return {
        "required_jobs": POLICY_JOBS,
        "events": {"pull_request": {"accepted_skips": ["full"]}, "push": {"accepted_skips": []}},
    }


def all_success(needs_overrides: dict | None = None) -> dict:
    needs = {job: "success" for job in POLICY_JOBS}
    needs.update(needs_overrides or {})
    return needs


def real_shape(results: dict) -> dict:
    """The shape ``toJSON(needs)`` really emits: one object per job."""
    return {job: {"result": result, "outputs": {}} for job, result in results.items()}


class VerdictTests(unittest.TestCase):
    """Every doubt must be a non-zero exit; success or declared skips exit 0."""

    def test_failure_conclusion_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), all_success({"unit": "failure"}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("root-cause failures", proc.stdout)
        self.assertIn("unit (failure)", proc.stdout)

    def test_cancelled_conclusion_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), all_success({"unit": "cancelled"}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("unit (cancelled)", proc.stdout)

    def test_known_job_absent_from_needs_exits_nonzero(self) -> None:
        needs = all_success()
        del needs["smoke"]
        proc = run_gate(base_policy(), needs)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("required job 'smoke' is absent from needs", proc.stdout)

    def test_undeclared_skip_for_event_exits_nonzero(self) -> None:
        # 'unit' is skipped but only 'full' is declared as an accepted skip.
        proc = run_gate(base_policy(), all_success({"unit": "skipped"}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("not declared for this event", proc.stdout)

    def test_undeclared_event_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), all_success(), event="workflow_dispatch")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("event 'workflow_dispatch' is not declared", proc.stdout)

    def test_unknown_needs_key_not_covered_by_policy_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), all_success({"surprise": "success"}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("needs key 'surprise' is not covered by the policy", proc.stdout)

    def test_unreadable_needs_file_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), needs_file="/nonexistent/needs.json")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("unreadable needs input", proc.stderr)

    def test_malformed_needs_json_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), stdin_text="{not json")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("not valid JSON", proc.stderr)

    def test_empty_needs_input_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), stdin_text="   \n")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("needs input is empty", proc.stderr)

    def test_all_success_exits_zero(self) -> None:
        proc = run_gate(base_policy(), all_success())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("VERDICT: PASS", proc.stdout)

    def test_declared_skips_exit_zero(self) -> None:
        proc = run_gate(base_policy(), all_success({"full": "skipped"}))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("declared skips (accepted)", proc.stdout)

    def test_output_separates_root_cause_from_cascade_skips(self) -> None:
        needs = all_success({"unit": "failure", "full": "skipped"})
        # push declares no accepted skips, so full lands in the cascade section.
        proc = run_gate(base_policy(), needs, event="push")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        root = proc.stdout.index("root-cause failures")
        cascade = proc.stdout.index("not declared for this event")
        self.assertIn("unit (failure)", proc.stdout[root:cascade])
        self.assertIn("full (skipped)", proc.stdout[cascade:])


class NeedsShapeTests(unittest.TestCase):
    """``toJSON(needs)`` emits ``{job: {result, outputs}}``; the flat
    ``{job: conclusion}`` shape stays accepted. Anything else is a doubt."""

    def test_real_needs_shape_all_success_exits_zero(self) -> None:
        proc = run_gate(base_policy(), real_shape(all_success()))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("VERDICT: PASS", proc.stdout)

    def test_real_needs_shape_as_github_serializes_it_exits_zero(self) -> None:
        # Byte-for-byte what the runner hands over: pretty-printed, multi-line.
        needs = real_shape(all_success({"full": "skipped"}))
        needs["lint"]["outputs"] = {"image": "it's quoted: 'x' and \"y\""}
        proc = run_gate(base_policy(), stdin_text=json.dumps(needs, indent=2))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("declared skips (accepted)", proc.stdout)

    def test_real_needs_shape_failure_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), real_shape(all_success({"unit": "failure"})))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("unit (failure)", proc.stdout)

    def test_real_needs_shape_undeclared_skip_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), real_shape(all_success({"unit": "skipped"})))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("not declared for this event", proc.stdout)

    def test_mixed_flat_and_object_values_are_read(self) -> None:
        needs = all_success()
        needs["unit"] = {"result": "cancelled", "outputs": {}}
        proc = run_gate(base_policy(), needs)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("unit (cancelled)", proc.stdout)

    def test_object_without_result_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), all_success({"unit": {"outputs": {}}}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("job 'unit'", proc.stderr)
        self.assertIn("'result'", proc.stderr)
        self.assertNotIn("Traceback", proc.stderr)

    def test_object_with_non_string_result_exits_nonzero(self) -> None:
        proc = run_gate(base_policy(), all_success({"unit": {"result": True, "outputs": {}}}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("job 'unit'", proc.stderr)
        self.assertNotIn("Traceback", proc.stderr)

    def test_value_of_any_other_type_exits_nonzero(self) -> None:
        for value in (None, 0, ["success"], True):
            with self.subTest(value=value):
                proc = run_gate(base_policy(), all_success({"unit": value}))
                self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
                self.assertIn("needs input must map", proc.stderr)
                self.assertNotIn("Traceback", proc.stderr)


class PolicyValidationTests(unittest.TestCase):
    """A malformed policy is a fail-closed message, never a traceback."""

    def assert_fail_closed(self, proc: subprocess.CompletedProcess, fragment: str) -> None:
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("FAIL fail-closed:", proc.stderr)
        self.assertIn(fragment, proc.stderr)
        self.assertNotIn("Traceback", proc.stderr)

    def test_event_without_accepted_skips_exits_nonzero(self) -> None:
        policy = base_policy()
        policy["events"]["pull_request"] = {}
        self.assert_fail_closed(run_gate(policy, all_success()), "accepted_skips")

    def test_event_without_accepted_skips_fails_even_for_another_event(self) -> None:
        # Validation must not depend on which event happens to be running.
        policy = base_policy()
        policy["events"]["schedule"] = {}
        self.assert_fail_closed(run_gate(policy, all_success(), event="push"), "event 'schedule'")

    def test_event_with_unknown_key_exits_nonzero(self) -> None:
        policy = base_policy()
        policy["events"]["push"] = {"accepted_skips": [], "allow_failures": ["unit"]}
        self.assert_fail_closed(run_gate(policy, all_success()), "event 'push'")

    def test_policy_that_is_not_utf8_exits_nonzero(self) -> None:
        proc = run_gate({}, all_success(), policy_bytes=b"\xff\xfe{}")
        self.assert_fail_closed(proc, "policy file")

    def test_needs_that_are_not_utf8_exit_nonzero(self) -> None:
        proc = run_gate(base_policy(), stdin_bytes=b"\xff\xfe{}")
        self.assert_fail_closed(proc, "needs input")

    def test_policy_nested_beyond_the_parser_limit_exits_nonzero(self) -> None:
        proc = run_gate({}, all_success(), policy_bytes=b"[" * 200_000)
        self.assert_fail_closed(proc, "policy file")

    def test_needs_nested_beyond_the_parser_limit_exit_nonzero(self) -> None:
        proc = run_gate(base_policy(), stdin_bytes=b"[" * 200_000)
        self.assert_fail_closed(proc, "needs input")

    def test_job_name_that_cannot_be_encoded_still_gets_a_verdict(self) -> None:
        # A lone surrogate is valid JSON and cannot be written as UTF-8.
        needs = '{"lint": "success", "\\ud800": "success"}'
        proc = run_gate(base_policy(), stdin_text=needs)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("VERDICT: FAIL", proc.stdout)
        self.assertIn("is not covered by the policy", proc.stdout)
        self.assertNotIn("Traceback", proc.stderr)

    def test_policy_of_unhashable_entries_exits_nonzero(self) -> None:
        policy = base_policy()
        policy["required_jobs"] = [["lint"], {"compile": 1}]
        self.assert_fail_closed(run_gate(policy, all_success()), "required_jobs")


class ExamplePolicyTests(unittest.TestCase):
    """The shipped example policy must itself satisfy the gate's contract."""

    def test_example_policy_is_valid(self) -> None:
        proc = subprocess.run(
            [sys.executable, str(SCRIPT), "--policy", str(EXAMPLE_POLICY),
             "--event", "push", "--needs-file", str(EXAMPLE_POLICY)],
            capture_output=True, text=True,
        )
        self.assertEqual(proc.returncode, 1)  # policy file is not a needs object: fail-closed
        self.assertIn("needs input must map", proc.stderr)


class RelaxationContractTests(unittest.TestCase):
    """El campo opcional 'relaxation' se valida siempre que esté presente."""

    def test_well_formed_relaxation_field_is_accepted(self) -> None:
        policy = base_policy()
        policy["relaxation"] = {"reason": "runner quota", "reference": "docs/ops#12", "issue": 135}
        proc = run_gate(policy, all_success())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_relaxation_missing_a_key_exits_nonzero(self) -> None:
        policy = base_policy()
        policy["relaxation"] = {"reason": "runner quota"}
        proc = run_gate(policy, all_success())
        self.assertNotEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("relaxation must be an object", proc.stderr)

    def test_relaxation_with_blank_value_exits_nonzero(self) -> None:
        policy = base_policy()
        policy["relaxation"] = {"reason": "   ", "reference": "docs/ops#12"}
        proc = run_gate(policy, all_success())
        self.assertNotEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("relaxation must be an object", proc.stderr)

    def test_relaxation_with_extra_key_exits_nonzero(self) -> None:
        policy = base_policy()
        policy["relaxation"] = {"reason": "r", "reference": "d", "ticket": "X"}
        proc = run_gate(policy, all_success())
        self.assertNotEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("relaxation must be an object", proc.stderr)


class VacuousGateTests(unittest.TestCase):
    """HR-15: un gate que no puede bloquear nada no es un PASS (#161)."""

    def test_accepted_skips_covering_every_required_job_fails_closed(self) -> None:
        # accepted_skips que cubre TODOS los jobs requeridos del evento:
        # el gate no podría bloquear nada, hoy pasa en verde.
        policy = base_policy()
        policy["events"]["pull_request"]["accepted_skips"] = list(POLICY_JOBS)
        proc = run_gate(policy, all_success())
        self.assertNotEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("accepted_skips covers every required job", proc.stdout + proc.stderr)

    def test_run_without_any_success_is_not_a_pass(self) -> None:
        # Un run donde ningún job requerido terminó en success (todos
        # skip-aceptados) no puede ser un PASS.
        policy = base_policy()
        policy["events"]["pull_request"]["accepted_skips"] = list(POLICY_JOBS)
        needs = {job: "skipped" for job in POLICY_JOBS}
        proc = run_gate(policy, needs)
        self.assertNotEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("no required job reported success", proc.stdout + proc.stderr)




class RelaxationIdentityTests(unittest.TestCase):
    """HR-33: una exención relaxation exige identidad verificable (actor o issue)."""
    def test_relaxation_without_identity_fails_closed(self) -> None:
        proc = run_gate(
            {**base_policy(),
             "relaxation": {"reason": "dropped", "reference": "doc"}},
            needs=real_shape(all_success()))
        self.assertNotEqual(proc.returncode, 0, proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
