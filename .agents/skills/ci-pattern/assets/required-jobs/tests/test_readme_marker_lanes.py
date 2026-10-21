#!/usr/bin/env python3
# ci-pattern asset tests — the marker_lanes example in the README really gates (#973)
"""The README documents the ``ui_changed`` lane of ``marker_lanes``; this suite
lifts that JSON example out of the README and runs it through the gate, so the
documentation cannot drift from the behaviour. Stdlib only."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
README = ASSET / "README.md"
SCRIPT = ASSET / "check_required_jobs.py"

JOBS = ["pr-gate", "lint", "typecheck", "e2e"]


def documented_policy() -> dict:
    """Return the first ```json fence of the README that declares marker_lanes."""
    for fence in README.read_text(encoding="utf-8").split("```json")[1:]:
        block = fence.split("```", 1)[0]
        if "marker_lanes" in block:
            return json.loads(block)
    raise AssertionError("README has no ```json example declaring marker_lanes")


def run_gate(policy: dict, ui_changed: str, skipped: tuple[str, ...]):
    needs = {job: {"result": "skipped" if job in skipped else "success",
                   "outputs": {"ui_changed": ui_changed} if job == "pr-gate" else {}}
             for job in JOBS}
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "policy.json"
        path.write_text(json.dumps(policy), encoding="utf-8")
        return subprocess.run(
            [sys.executable, str(SCRIPT), "--policy", str(path), "--event", "pull_request"],
            input=json.dumps(needs), capture_output=True, text=True, encoding="utf-8")


class ReadmeMarkerLaneExample(unittest.TestCase):

    def test_documents_marker_lanes_and_equals(self) -> None:
        text = README.read_text(encoding="utf-8")
        self.assertIn("| `marker_lanes` |", text)
        self.assertIn("`equals`", text)

    def test_example_accepts_e2e_skip_without_ui_changes(self) -> None:
        r = run_gate(documented_policy(), "false", ("e2e",))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("VERDICT: PASS", r.stdout)

    def test_example_rejects_e2e_skip_with_ui_changes(self) -> None:
        r = run_gate(documented_policy(), "true", ("e2e",))
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        self.assertIn("not declared for this event", r.stdout)

    def test_example_accepts_no_other_skip(self) -> None:
        r = run_gate(documented_policy(), "false", ("e2e", "typecheck"))
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        self.assertIn("not declared for this event", r.stdout)


if __name__ == "__main__":
    unittest.main(defaultTest="ReadmeMarkerLaneExample")
