#!/usr/bin/env python3
# ci-pattern asset tests — docs-only lane skips accepted only on a proven docs_changed marker (DysTelefonica/team-skills#857)
"""Executable suite for the docs-only lane of the required-jobs aggregator.

The lane declares its event, the gate job that publishes the marker, the
output key that holds it and the skips it covers. The aggregator accepts
those skips only when the policy declares the lane and the marker is
exactly ``"true"`` on that event; any other value, a missing output, or a
different event keeps every skip undeclared and the gate fails closed.
The lane coexists with the optional ``dependencies`` map (#853) — both
keys cover disjoint skips and the gate accepts the union.

Runs with the stdlib only, the same way CI invokes it:
``python3 personal/ardelperal/ci-pattern/assets/required-jobs/tests/test_docs_only_lane.py``.
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

UNDECLARED = "not declared for this event"
DOCS_SKIPS = ["typecheck", "test", "integration", "e2e"]

POLICY = {
    "required_jobs": ["pr-gate", "lint", "typecheck", "test", "integration", "e2e"],
    "events": {
        "pull_request": {"accepted_skips": []},
        "push": {"accepted_skips": []},
    },
    "docs_only_lane": {
        "event": "pull_request",
        "gate_job": "pr-gate",
        "output": "docs_changed",
        "skips": DOCS_SKIPS,
    },
}


def docs_needs(marker: object = "true", skipped: list[str] | None = None) -> dict:
    """A docs-only pull_request: pr-gate green with the marker, heavy jobs skipped."""
    skipped = DOCS_SKIPS if skipped is None else skipped
    needs = {"pr-gate": {"result": "success", "outputs": {"docs_changed": marker}}}
    for job in ["lint", *DOCS_SKIPS]:
        needs[job] = {"result": "skipped" if job in skipped else "success", "outputs": {}}
    return needs


def run_gate(policy: dict, needs: dict, event: str = "pull_request",
             base: dict | None = None) -> subprocess.CompletedProcess:
    """Run the gate as CI does, with the needs JSON on stdin."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        policy_path = tmp_path / "policy.json"
        policy_path.write_text(json.dumps(policy), encoding="utf-8")
        argv = [sys.executable, str(SCRIPT), "--policy", str(policy_path), "--event", event]
        if base is not None:
            base_path = tmp_path / "base.json"
            base_path.write_text(json.dumps(base), encoding="utf-8")
            argv += ["--base-policy", str(base_path)]
        return subprocess.run(argv, input=json.dumps(needs), capture_output=True,
                              text=True, encoding="utf-8", check=False)


class DocsOnlyLaneTest(unittest.TestCase):
    def test_docs_only_pull_request_with_declared_lane_passes(self) -> None:
        result = run_gate(POLICY, docs_needs())
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_same_skips_without_the_marker_fail(self) -> None:
        result = run_gate(POLICY, docs_needs(marker="false"))
        self.assertEqual(result.returncode, 1)
        self.assertIn(UNDECLARED, result.stdout)

    def test_marker_absent_fails(self) -> None:
        needs = docs_needs()
        needs["pr-gate"]["outputs"] = {}
        result = run_gate(POLICY, needs)
        self.assertEqual(result.returncode, 1)
        self.assertIn(UNDECLARED, result.stdout)

    def test_marker_must_be_exactly_true(self) -> None:
        result = run_gate(POLICY, docs_needs(marker=True))
        self.assertEqual(result.returncode, 1)
        self.assertIn(UNDECLARED, result.stdout)

    def test_marker_on_push_does_not_open_the_lane(self) -> None:
        result = run_gate(POLICY, docs_needs(), event="push")
        self.assertEqual(result.returncode, 1)
        self.assertIn(UNDECLARED, result.stdout)

    def test_skip_outside_the_lane_list_fails(self) -> None:
        result = run_gate(POLICY, docs_needs(skipped=DOCS_SKIPS + ["lint"]))
        self.assertEqual(result.returncode, 1)
        self.assertIn("lint", result.stdout)

    def test_root_cause_still_fails_on_a_docs_only_pull_request(self) -> None:
        needs = docs_needs()
        needs["test"] = {"result": "failure", "outputs": {}}
        result = run_gate(POLICY, needs)
        self.assertEqual(result.returncode, 1)
        self.assertIn("test (failure)", result.stdout)

    def test_policy_without_the_lane_keeps_the_marker_inert(self) -> None:
        policy = copy.deepcopy(POLICY)
        del policy["docs_only_lane"]
        result = run_gate(policy, docs_needs())
        self.assertEqual(result.returncode, 1)
        self.assertIn(UNDECLARED, result.stdout)

    def test_lane_gate_job_outside_required_jobs_exits_nonzero(self) -> None:
        policy = copy.deepcopy(POLICY)
        policy["docs_only_lane"]["gate_job"] = "ghost"
        result = run_gate(policy, docs_needs())
        self.assertEqual(result.returncode, 1)
        self.assertIn("docs_only_lane", result.stderr)

    def test_lane_skip_outside_required_jobs_exits_nonzero(self) -> None:
        policy = copy.deepcopy(POLICY)
        policy["docs_only_lane"]["skips"] = DOCS_SKIPS + ["ghost"]
        result = run_gate(policy, docs_needs())
        self.assertEqual(result.returncode, 1)
        self.assertIn("docs_only_lane", result.stderr)

    def test_lane_with_unknown_key_exits_nonzero(self) -> None:
        policy = copy.deepcopy(POLICY)
        policy["docs_only_lane"]["unconditional"] = True
        result = run_gate(policy, docs_needs())
        self.assertEqual(result.returncode, 1)
        self.assertIn("docs_only_lane", result.stderr)

    def test_adding_the_lane_without_a_relaxation_field_exits_nonzero(self) -> None:
        base = copy.deepcopy(POLICY)
        del base["docs_only_lane"]
        result = run_gate(POLICY, docs_needs(), base=base)
        self.assertEqual(result.returncode, 1)
        self.assertIn("docs-only skip", result.stdout)

    def test_adding_the_lane_with_its_relaxation_field_passes(self) -> None:
        base = copy.deepcopy(POLICY)
        del base["docs_only_lane"]
        candidate = copy.deepcopy(POLICY)
        candidate["relaxation"] = {
            "reason": "docs-only pull_request skips typecheck, test, integration and e2e "
                      "when pr-gate proves docs_changed=true (#857, after APAP_WEB#2177)",
            "reference": "DysTelefonica/team-skills#857; APAP_WEB#1196, APAP_WEB#2177",
            "issue": 857,
        }
        result = run_gate(candidate, docs_needs(), base=base)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_lane_coexists_with_dependencies_and_covers_disjoint_skips(self) -> None:
        """The lane and the dependencies map (#853) are both optional and combinable.

        ``build`` is in ``accepted_skips.pull_request`` and ``e2e`` depends on it;
        ``typecheck/test/integration`` are covered by the docs-only lane. A
        docs-only PR skips ``build``, ``e2e`` and the three lane jobs: the
        cascade from dependencies accepts ``e2e`` once build is declared, and
        the lane also accepts its skips. The gate accepts the union of both
        without widening ``accepted_skips``.
        """
        policy = copy.deepcopy(POLICY)
        policy["required_jobs"] = ["pr-gate", "lint", "build", "e2e",
                                    "typecheck", "test", "integration"]
        policy["events"]["pull_request"]["accepted_skips"] = ["build"]
        policy["dependencies"] = {"e2e": ["build"]}
        policy["docs_only_lane"]["skips"] = ["typecheck", "test", "integration"]
        needs = {
            "pr-gate": {"result": "success", "outputs": {"docs_changed": "true"}},
            "lint": {"result": "success", "outputs": {}},
            "build": {"result": "skipped", "outputs": {}},
            "e2e": {"result": "skipped", "outputs": {}},
            "typecheck": {"result": "skipped", "outputs": {}},
            "test": {"result": "skipped", "outputs": {}},
            "integration": {"result": "skipped", "outputs": {}},
        }
        result = run_gate(policy, needs)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("VERDICT: PASS", result.stdout)

    def test_lane_alone_is_silent_when_dependency_cascade_would_save_a_skip(self) -> None:
        """The lane does NOT widen the cascade of dependencies.

        ``build`` is skipped and is in ``accepted_skips.pull_request``; ``e2e``
        depends on ``build``. Without the lane covering ``e2e`` (only
        ``typecheck/test/integration`` are listed), the cascade from
        dependencies still rescues ``e2e`` — and the lane's three skips cover
        the heavy docs jobs the cascade doesn't reach.
        """
        policy = copy.deepcopy(POLICY)
        policy["required_jobs"] = ["pr-gate", "lint", "build", "e2e",
                                    "typecheck", "test", "integration"]
        policy["events"]["pull_request"]["accepted_skips"] = ["build"]
        policy["dependencies"] = {"e2e": ["build"]}
        policy["docs_only_lane"]["skips"] = ["typecheck", "test", "integration"]
        needs = {
            "pr-gate": {"result": "success", "outputs": {"docs_changed": "true"}},
            "lint": {"result": "success", "outputs": {}},
            "build": {"result": "skipped", "outputs": {}},
            "e2e": {"result": "skipped", "outputs": {}},
            "typecheck": {"result": "skipped", "outputs": {}},
            "test": {"result": "skipped", "outputs": {}},
            "integration": {"result": "skipped", "outputs": {}},
        }
        result = run_gate(policy, needs)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("VERDICT: PASS", result.stdout)

    def test_dependencies_cascade_alone_still_works_without_the_lane(self) -> None:
        """When the policy declares ``dependencies`` but not ``docs_only_lane``,
        the cascade from #853 still accepts transitive skips. The lane is
        optional; its absence must not disable dependencies."""
        policy = copy.deepcopy(POLICY)
        del policy["docs_only_lane"]
        policy["required_jobs"] = ["pr-gate", "lint", "build", "e2e"]
        policy["events"]["pull_request"]["accepted_skips"] = ["build"]
        policy["dependencies"] = {"e2e": ["build"]}
        # Marker still true: the gate ignores it without the lane.
        needs = {
            "pr-gate": {"result": "success", "outputs": {"docs_changed": "true"}},
            "lint": {"result": "success", "outputs": {}},
            "build": {"result": "skipped", "outputs": {}},
            "e2e": {"result": "skipped", "outputs": {}},
        }
        result = run_gate(policy, needs)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("VERDICT: PASS", result.stdout)

    def test_lane_event_keyword_must_be_a_string(self) -> None:
        """The marker must be a string equal to ``\"true\"``; non-string outputs
        from the gate do not open the lane, even when truthy."""
        result = run_gate(POLICY, docs_needs(marker="True"))  # capital T
        self.assertEqual(result.returncode, 1)
        self.assertIn(UNDECLARED, result.stdout)


if __name__ == "__main__":
    unittest.main()
