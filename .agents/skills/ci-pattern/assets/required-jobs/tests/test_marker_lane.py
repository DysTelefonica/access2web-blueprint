#!/usr/bin/env python3
# ci-pattern asset tests — ``marker_lanes`` list (#863). Run with stdlib.
"""Executable suite for the marker-lanes list of the required-jobs gate."""

from __future__ import annotations

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
UI_SKIPS = ["e2e"]


def _run(policy, needs, base=None):
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp) / "p.json"
        p.write_text(json.dumps(policy), encoding="utf-8")
        argv = [sys.executable, str(SCRIPT), "--policy", str(p),
                "--event", "pull_request"]
        if base is not None:
            (Path(tmp) / "b.json").write_text(json.dumps(base), encoding="utf-8")
            argv += ["--base-policy", str(Path(tmp) / "b.json")]
        return subprocess.run(argv, input=json.dumps(needs),
                              capture_output=True, text=True, encoding="utf-8")


def _needs(docs="false", ui="true", *, skip=()):
    out = {"pr-gate": {"result": "success",
                       "outputs": {"docs_changed": docs, "ui_changed": ui}},
           "lint": {"result": "success", "outputs": {}}}
    for j in DOCS_SKIPS:
        out[j] = {"result": "skipped" if j in skip else "success",
                   "outputs": {}}
    return out


def _base():
    return {"required_jobs": ["pr-gate", "lint", *DOCS_SKIPS],
            "events": {"pull_request": {"accepted_skips": []},
                       "push": {"accepted_skips": []}}}


def _pair():
    """APAP_WEB-shaped pair: docs (default) + ui (equals=false)."""
    return [{"event": "pull_request", "gate_job": "pr-gate",
             "output": "docs_changed", "skips": DOCS_SKIPS},
            {"event": "pull_request", "gate_job": "pr-gate",
             "output": "ui_changed", "equals": "false", "skips": UI_SKIPS}]


class TwoSimultaneous(unittest.TestCase):

    def test_combinations(self) -> None:
        # (docs, ui, expected_rc, expected_text) — heavy jobs are skipped.
        cases = [("true", "false", 0, "VERDICT: PASS"),    # both open
                 ("true", "true", 0, "VERDICT: PASS"),     # only docs
                 ("false", "true", 1, UNDECLARED)]         # neither
        for docs, ui, rc, text in cases:
            with self.subTest(docs=docs, ui=ui):
                policy = {**_base(), "marker_lanes": _pair()}
                r = _run(policy, _needs(docs, ui, skip=DOCS_SKIPS))
                self.assertEqual(r.returncode, rc, r.stdout + r.stderr)
                self.assertIn(text, r.stdout)

    def test_only_ui_with_only_e2e_skipped(self) -> None:
        # APAP_WEB shape: docs=false, ui=false, e2e skipped; heavy
        # tests still pass.
        policy = {**_base(), "marker_lanes": _pair()}
        r = _run(policy, _needs("false", "false", skip=("e2e",)))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("VERDICT: PASS", r.stdout)

    def test_compat_with_docs_only_lane_unchanged(self) -> None:
        # Pre-#863 shape still works: docs_only_lane alone is treated
        # as a single-entry marker lane.
        policy = {**_base(), "docs_only_lane": {
            "event": "pull_request", "gate_job": "pr-gate",
            "output": "docs_changed", "skips": DOCS_SKIPS}}
        r = _run(policy, _needs("true", skip=DOCS_SKIPS))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("VERDICT: PASS", r.stdout)


class Validation(unittest.TestCase):

    def test_load_failures(self) -> None:
        lane = {"event": "pull_request", "gate_job": "pr-gate",
                "output": "docs_changed", "skips": ["e2e"]}
        cases = [("empty list", {**_base(), "marker_lanes": []}),
                 ("duplicate signature within marker_lanes",
                  {**_base(), "marker_lanes": [lane, lane]}),
                 ("equals invalid value",
                  {**_base(), "marker_lanes": [
                      {"event": "pull_request", "gate_job": "pr-gate",
                       "output": "ui_changed", "equals": None,
                       "skips": UI_SKIPS}]}),
                 ("ghost skip",
                  {**_base(), "marker_lanes": [
                      {"event": "pull_request", "gate_job": "pr-gate",
                       "output": "ui_changed", "skips": ["ghost"]}]})]
        for label, policy in cases:
            with self.subTest(case=label):
                r = _run(policy, _needs())
                self.assertEqual(r.returncode, 1)
                self.assertIn("marker_lanes", r.stderr)

    def test_cross_source_duplicate_with_docs_only_lane_rejected(self) -> None:
        # A marker_lanes entry sharing the same
        # ``(event, gate_job, output)`` triple as the legacy
        # ``docs_only_lane`` is rejected at load time (#863).
        policy = {**_base(),
                  "docs_only_lane": {
                      "event": "pull_request", "gate_job": "pr-gate",
                      "output": "ui_changed", "skips": UI_SKIPS},
                  "marker_lanes": [
                      {"event": "pull_request", "gate_job": "pr-gate",
                       "output": "ui_changed", "equals": "false",
                       "skips": ["e2e"]}]}
        r = _run(policy, _needs("false", "false", skip=("e2e",)))
        self.assertEqual(r.returncode, 1)
        self.assertIn("duplicates docs_only_lane", r.stderr)
        self.assertIn("ui_changed", r.stderr)

    def test_distinct_sources_accepted(self) -> None:
        # Sanity: when docs_only_lane and marker_lanes have different
        # ``(event, gate_job, output)`` triples, both are accepted.
        policy = {**_base(),
                  "docs_only_lane": {
                      "event": "pull_request", "gate_job": "pr-gate",
                      "output": "docs_changed", "skips": DOCS_SKIPS},
                  "marker_lanes": [
                      {"event": "pull_request", "gate_job": "pr-gate",
                       "output": "ui_changed", "equals": "false",
                       "skips": UI_SKIPS}]}
        r = _run(policy, _needs("true", "false", skip=DOCS_SKIPS))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("VERDICT: PASS", r.stdout)


class Coexistence(unittest.TestCase):
    # build in accepted_skips; e2e cascades (#853); the marker_lane
    # covers e2e via ui_changed=false. Union accepted.
    def test_marker_lane_with_dependencies(self) -> None:
        policy = _base()
        policy["required_jobs"].append("build")
        policy["events"]["pull_request"]["accepted_skips"] = ["build"]
        policy["dependencies"] = {"e2e": ["build"]}
        policy["marker_lanes"] = [
            {"event": "pull_request", "gate_job": "pr-gate",
             "output": "ui_changed", "equals": "false", "skips": ["e2e"]}]
        needs = {"pr-gate": {"result": "success",
                             "outputs": {"ui_changed": "false"}},
                 "lint": {"result": "success", "outputs": {}},
                 "typecheck": {"result": "success", "outputs": {}},
                 "test": {"result": "success", "outputs": {}},
                 "integration": {"result": "success", "outputs": {}},
                 "build": {"result": "skipped", "outputs": {}},
                 "e2e": {"result": "skipped", "outputs": {}}}
        r = _run(policy, needs)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("VERDICT: PASS", r.stdout)


class Relaxations(unittest.TestCase):

    def _passing(self):
        return {"pr-gate": {"result": "success",
                             "outputs": {"ui_changed": "false",
                                         "docs_changed": "true"}},
                "lint": {"result": "success", "outputs": {}},
                "typecheck": {"result": "success", "outputs": {}},
                "test": {"result": "success", "outputs": {}},
                "integration": {"result": "success", "outputs": {}},
                "e2e": {"result": "success", "outputs": {}}}

    def test_relaxation_cases(self) -> None:
        ui_lane = {"event": "pull_request", "gate_job": "pr-gate",
                   "output": "ui_changed", "equals": "false",
                   "skips": UI_SKIPS}
        cases = [
            # (label, base, candidate, expected_text, expects_refused)
            ("new lane without relaxation is refused",
             _base(), {**_base(), "marker_lanes": [ui_lane]},
             "marker lane added", True),
            ("new lane with relaxation passes",
             _base(),
             {**_base(), "marker_lanes": [ui_lane],
              "relaxation": {
                  "reason": "marker-lane #863: ui_only_lane opens on ui_changed=false",
                  "reference": "DysTelefonica/team-skills#863", "issue": 863}},
             "marker lane added", False),
            ("skip added to existing lane",
             {**_base(), "marker_lanes": [
                 {"event": "pull_request", "gate_job": "pr-gate",
                  "output": "docs_changed",
                  "skips": ["typecheck", "test"]}]},
             {**_base(), "marker_lanes": [
                 {"event": "pull_request", "gate_job": "pr-gate",
                  "output": "docs_changed",
                  "skips": ["typecheck", "test", "integration"]}]},
             "marker-lane skip 'integration' added", True),
            ("source change",
             {**_base(), "marker_lanes": [
                 {"event": "pull_request", "gate_job": "pr-gate",
                  "output": "ui_changed", "skips": UI_SKIPS}]},
             {**_base(), "marker_lanes": [ui_lane]},
             "marker-lane source changed", True),
        ]
        for label, base, candidate, expected_text, refused in cases:
            with self.subTest(case=label):
                r = _run(candidate, self._passing(), base=base)
                self.assertIn(expected_text, r.stdout,
                               r.stdout + r.stderr)
                if refused:
                    self.assertIn("REFUSED", r.stdout)
                else:
                    self.assertEqual(r.returncode, 0,
                                      r.stdout + r.stderr)


class MarkerContract(unittest.TestCase):
    # The marker comparison is string-exact; Python ``False`` does not
    # equal the string ``"false"`` and the lane does not open.
    def test_python_bool_does_not_open_lane(self) -> None:
        policy = {**_base(), "marker_lanes": [
            {"event": "pull_request", "gate_job": "pr-gate",
             "output": "ui_changed", "equals": "false", "skips": UI_SKIPS}]}
        needs = _needs("false", ui="ignored", skip=("e2e",))
        needs["pr-gate"]["outputs"]["ui_changed"] = False
        r = _run(policy, needs)
        self.assertEqual(r.returncode, 1)
        self.assertIn(UNDECLARED, r.stdout)


__all__ = ["TwoSimultaneous", "Validation", "Coexistence",
            "Relaxations", "MarkerContract"]


if __name__ == "__main__":
    unittest.main()
