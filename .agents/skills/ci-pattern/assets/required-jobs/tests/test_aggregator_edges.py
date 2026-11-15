#!/usr/bin/env python3
# ci-pattern asset tests — aggregator edges and cause naming (DysTelefonica/team-skills#1208, phase 2 of #1179)
"""Edge cases of the required-jobs aggregator: exact AND, the conclusions the
runner can emit besides success/failure, re-runs, and a cause that names the
red check instead of only the job.

One named test per edge of #1208; the edges that already held before #1208
are pinned here too, so a later change cannot silently drop them.
"""

from __future__ import annotations

import importlib.util
import json
import random
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "check_required_jobs.py"

# The 12 required jobs of DysTelefonica/APAP_WEB (required-jobs-policy.json). Using the real set
# keeps the edge cases on the same surface the phase-2 gate will judge.
REQUIRED = [
    "build", "docs", "e2e", "integration", "lint", "mutation", "pr-gate",
    "security", "security-deep", "security-on-deps", "test", "typecheck",
]
# Accepted skips on pull_request for that repo; every other skip is undeclared.
PR_ACCEPTED = ["build", "mutation", "security-deep"]
RED_CONCLUSIONS = ["failure", "cancelled", "timed_out", "startup_failure"]


def _policy() -> dict:
    return {
        "required_jobs": REQUIRED,
        "events": {
            "pull_request": {"accepted_skips": PR_ACCEPTED},
            "push": {"accepted_skips": []},
        },
    }


def _needs(results: dict) -> dict:
    """The shape ``toJSON(needs)`` emits: one object per job."""
    return {job: {"result": result, "outputs": {}} for job, result in results.items()}


def _all_success() -> dict:
    return {job: "success" for job in REQUIRED}


def _run(needs: dict, *, extra: list[str] | None = None, policy: dict | None = None,
         event: str = "pull_request") -> subprocess.CompletedProcess:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        policy_path = tmp_path / "policy.json"
        policy_path.write_text(json.dumps(policy or _policy()), encoding="utf-8")
        needs_path = tmp_path / "needs.json"
        needs_path.write_text(json.dumps(_needs(needs)), encoding="utf-8")
        argv = [sys.executable, str(SCRIPT), "--policy", str(policy_path), "--event", event,
                "--needs-file", str(needs_path), *(extra or [])]
        proc = subprocess.run(argv, capture_output=True, text=True, encoding="utf-8")
        return proc


def _load_checker():
    spec = importlib.util.spec_from_file_location("check_required_jobs", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class ExactAndTests(unittest.TestCase):
    """(a): the verdict is green iff every required job is success or an accepted skip."""

    def test_all_success_exits_zero(self) -> None:
        self.assertEqual(_run(_all_success()).returncode, 0)

    def test_each_required_job_red_alone_exits_nonzero(self) -> None:
        # A2: one red conclusion per required job, for each conclusion the runner can emit.
        for job in REQUIRED:
            for conclusion in RED_CONCLUSIONS:
                with self.subTest(job=job, conclusion=conclusion):
                    needs = _all_success()
                    needs[job] = conclusion
                    proc = _run(needs)
                    self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
                    self.assertIn(f"{job} ({conclusion})", proc.stdout)

    def test_random_needs_agree_with_the_and_reference(self) -> None:
        # A3: 500 generated needs sets; the verdict must equal the AND written out by hand.
        checker = _load_checker()
        policy = _policy()
        rng = random.Random(1208)
        outcomes = ["success", "success", "success", "failure", "cancelled", "timed_out",
                    "startup_failure", "skipped"]
        for case in range(500):
            needs = {job: rng.choice(outcomes) for job in REQUIRED}
            accepted = set(PR_ACCEPTED)
            every_ok = all(
                conclusion == "success" or (conclusion == "skipped" and job in accepted)
                for job, conclusion in needs.items()
            )
            any_success = any(conclusion == "success" for conclusion in needs.values())
            # HR-15: un run sin ningun job verde tampoco es PASS. Con 3 de 12 skips
            # aceptados, `every_ok` ya lo implica: el termino solo lo deja explicito.
            reference = every_ok and any_success
            got = checker.evaluate(needs, policy, "pull_request")["passed"]
            self.assertEqual(got, reference, f"case {case}: {needs}")


class EdgeTests(unittest.TestCase):
    """(b): one named test per edge of #1208."""

    def test_accepted_skips_on_pull_request_exit_zero(self) -> None:
        # B1: build, mutation and security-deep are skipped on pull_request by design.
        needs = _all_success()
        for job in PR_ACCEPTED:
            needs[job] = "skipped"
        self.assertEqual(_run(needs).returncode, 0)

    def test_same_skips_on_push_exit_nonzero(self) -> None:
        # B2: the same skips are not accepted on push.
        needs = _all_success()
        needs["build"] = "skipped"
        self.assertEqual(_run(needs, event="push").returncode, 1)

    def test_cancelled_names_the_cause(self) -> None:
        # B3: exit 1 already held; the cause line must name the cancelled job.
        needs = _all_success()
        needs["lint"] = "cancelled"
        proc = _run(needs)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("blocked by: lint (cancelled)", proc.stdout)

    def test_required_job_absent_from_needs_exits_nonzero(self) -> None:
        # B4: a required job that never reached needs fails closed.
        needs = _all_success()
        del needs["test"]
        self.assertEqual(_run(needs).returncode, 1)

    def test_startup_failure_is_a_root_cause(self) -> None:
        # B5: a job that never started is a root cause, not an unknown conclusion.
        needs = _all_success()
        needs["test"] = "startup_failure"
        proc = _run(needs)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("blocked by: test (startup_failure)", proc.stdout)
        self.assertNotIn("unknown conclusion", proc.stdout)

    def test_timed_out_is_a_root_cause(self) -> None:
        # B6: a timeout is a root cause, named as such.
        needs = _all_success()
        needs["integration"] = "timed_out"
        proc = _run(needs)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("blocked by: integration (timed_out)", proc.stdout)
        self.assertNotIn("unknown conclusion", proc.stdout)

    def test_rerun_turns_red_into_green_judges_the_second_attempt(self) -> None:
        # B7: the needs passed are those of attempt 2, all green; the verdict says which attempt it judged.
        proc = _run(_all_success(), extra=["--run-attempt", "2"])
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("judged attempt 2", proc.stdout)
        self.assertNotIn("attempt 1", proc.stdout)

    def test_rerun_turns_green_into_red_judges_the_second_attempt(self) -> None:
        # B8: attempt 2 has a required failure, so the verdict is red and names attempt 2.
        needs = _all_success()
        needs["typecheck"] = "failure"
        proc = _run(needs, extra=["--run-attempt", "2"])
        self.assertEqual(proc.returncode, 1)
        self.assertIn("judged attempt 2", proc.stdout)
        self.assertIn("blocked by: typecheck (failure)", proc.stdout)


class CauseTests(unittest.TestCase):
    """(c): the cause names the red check, not only the job."""

    def test_cause_quotes_the_first_annotation_of_the_red_check(self) -> None:
        # C2: an annotation file maps the job to the first error the check reported.
        needs = _all_success()
        needs["test"] = "failure"
        with tempfile.TemporaryDirectory() as tmp:
            ann = Path(tmp) / "annotations.json"
            ann.write_text(json.dumps({"test": "pytest: 3 failed, tests/test_x.py::test_y"}),
                           encoding="utf-8")
            proc = _run(needs, extra=["--annotations", str(ann)])
        self.assertEqual(proc.returncode, 1)
        self.assertIn("test (failure): pytest: 3 failed, tests/test_x.py::test_y", proc.stdout)

    def test_red_check_without_annotation_says_so(self) -> None:
        # C3: no annotation for the red job is said explicitly, never silently dropped.
        needs = _all_success()
        needs["test"] = "failure"
        proc = _run(needs)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("test (failure): no annotation", proc.stdout)

    def test_cascade_skip_is_never_the_cause(self) -> None:
        # C1 (pinned): the cause lists the root failure, not the skips it cascaded
        # into; neither an undeclared skip nor a declared cascade is ever a cause.
        needs = _all_success()
        needs["lint"] = "failure"
        needs["security"] = "skipped"
        proc = _run(needs)
        self.assertIn("blocked by: lint (failure)", proc.stdout)
        self.assertNotIn("blocked by: security", proc.stdout)
        policy = _policy()
        policy["dependencies"] = {"security": ["build"]}
        needs = _all_success()
        needs["build"] = "skipped"
        needs["security"] = "skipped"
        proc = _run(needs, policy=policy)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("blocked by", proc.stdout)
        self.assertIn("declared skips (accepted)", proc.stdout)

    def test_cause_names_every_red_check_and_no_green_one(self) -> None:
        # (c): the cause names the red checks, with their conclusion; a green check
        # and an accepted skip never appear in the cause list.
        needs = _all_success()
        needs["test"] = "failure"
        needs["lint"] = "timed_out"
        needs["build"] = "skipped"
        proc = _run(needs)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("blocked by: lint (timed_out), test (failure)", proc.stdout)
        self.assertNotIn("blocked by: build", proc.stdout)


class InputContractTests(unittest.TestCase):
    """The new inputs are fail-closed: a doubt never becomes a half verdict."""

    def test_run_attempt_must_be_a_positive_integer(self) -> None:
        # --run-attempt 0 is not an attempt GitHub can issue; the gate doubts it.
        proc = _run(_all_success(), extra=["--run-attempt", "0"])
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("run attempt must be a positive integer", proc.stderr)

    def test_unusable_annotations_file_fails_closed(self) -> None:
        # An annotations file that is not a job -> message object, or that cannot
        # be read at all, is a doubt: the gate does not judge a medias.
        with tempfile.TemporaryDirectory() as tmp:
            ann = Path(tmp) / "annotations.json"
            ann.write_text(json.dumps({"test": {"message": "nested"}}), encoding="utf-8")
            malformed = _run(_all_success(), extra=["--annotations", str(ann)])
        unreadable = _run(_all_success(), extra=["--annotations", "/nonexistent/ann.json"])
        self.assertEqual(malformed.returncode, 1, malformed.stdout + malformed.stderr)
        self.assertIn("annotations file must be a JSON object of job -> message",
                      malformed.stderr)
        self.assertEqual(unreadable.returncode, 1, unreadable.stdout + unreadable.stderr)
        self.assertIn("unreadable annotations file", unreadable.stderr)


if __name__ == "__main__":
    unittest.main()
