#!/usr/bin/env python3
# ci-pattern asset tests — workflow policy gate, structural reader + HR-26 (DysTelefonica/team-skills#192)
"""Executable suite for the workflow policy gate (black-box, like CI).

The structural reader carries positive, negative and degenerate/empty-subject
cases (exit 2); HR-26 carries positive findings, negatives that must stay
clean, and degenerate jobs. Fail-closed parsing exits 2; findings exit 1;
clean exit 0. HR-3 and HR-30 cases join in the next chained PR (#192)."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_workflow_policy.py"
sys.path.insert(0, str(GATE.parent))
import check_workflow_policy as gate  # noqa: E402

POLICY = {"required_jobs": ["build"], "events": {"pull_request": {"accepted_skips": []}}}
PINNED = "uses: actions/checkout@11bd71901bbe5b1630ceea73d27597364c9af683"
RUN_OK = "set -u\n          exit \"$fail\""


def yml(run=RUN_OK, uses=PINNED, name="build (${{ github.event_name }})",
        on="on:\n  pull_request:\n", job="  build:", head="", tail=""):
    s = f"{on}\njobs:\n{head}{job}\n    name: {name}\n    steps:\n"
    if uses:
        s += f"      - {uses}\n"
    if run:
        s += f"        run: |\n          {run}\n"
    return s + tail


class GateCase(unittest.TestCase):
    def run_gate(self, text, policy=POLICY, extra=()):
        with tempfile.TemporaryDirectory() as d:
            w, p = Path(d, "wf.yml"), Path(d, "policy.json")
            w.write_text(text, encoding="utf-8")
            p.write_text(json.dumps(policy), encoding="utf-8")
            out = subprocess.run([sys.executable, str(GATE), "--workflow", str(w),
                                  "--required-jobs", str(p), "--base-branch", "main", *extra],
                                 capture_output=True, text=True)
        return out.returncode, out.stdout + out.stderr

    def table(self, cases):
        for label, text, code, needle in cases:
            with self.subTest(case=label):
                got, out = self.run_gate(text)
                self.assertEqual(got, code, out)
                if needle:
                    self.assertIn(needle, out)
                elif code == 0:
                    self.assertNotIn("HR-", out)


class InvalidInput(GateCase):
    """Exit 2: unreadable, malformed, unsupported, or empty subject."""
    CASES = [
        ("empty file", "", None),
        ("comments only", "# only a comment\n", None),
        ("no jobs block", "on:\n  pull_request:\n", None),
        ("empty jobs mapping", "on:\n  pull_request:\njobs: {}\n", None),
        ("tab indentation", "on:\n  pull_request:\njobs:\n\tbuild:\n", None),
        ("alias value", "on:\n  pull_request:\njobs:\n  build: *a\n", None),
        ("anchor value", "on:\n  pull_request:\njobs:\n  build: &a\n    steps: []\n", None),
        ("merge key", "on:\n  pull_request:\njobs:\n  <<: *base\n", None),
        ("jobs as block sequence", "on:\n  pull_request:\njobs:\n  - build\n", None),
    ]

    def test_exit_two(self):
        self.table([(l, t, 2, None) for l, t, _ in self.CASES])

    def test_invalid_policy(self):
        for label, pol in [("no required_jobs", {"events": {}}),
                           ("empty required_jobs", {"required_jobs": [], "events": {}})]:
            with self.subTest(case=label):
                code, out = self.run_gate(yml(), policy=pol)
                self.assertEqual(code, 2)
                self.assertIn("required_jobs", out)

    def test_nonexistent_workflow(self):
        self.assertEqual(self.run_gate(yml(), extra=("--workflow", "/no/such.yml"))[0], 2)

    def test_missing_required_argument(self):
        out = subprocess.run([sys.executable, str(GATE)], capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)


class HR26PinnedToolchain(GateCase):
    """HR-26: actions pinned by 40-hex SHA and exact tool versions, every job."""
    def test_findings(self):
        self.table([
            ("action tag", yml(uses="uses: actions/checkout@v4"), 1, "HR-26"),
            ("short sha", yml(uses="uses: actions/checkout@11bd719"), 1, "HR-26"),
            ("action without ref", yml(uses="uses: actions/checkout"), 1, "HR-26"),
            ("setup-python minor only", yml(
                tail="      - uses: actions/setup-python@" + "a"*40 +
                     "\n        with:\n          python-version: '3.12'\n"), 1, "HR-26"),
            ("setup-node major only", yml(
                tail="      - uses: actions/setup-node@" + "b"*40 +
                     "\n        with:\n          node-version: 20\n"), 1, "HR-26"),
            ("setup-node missing version", yml(
                tail="      - uses: actions/setup-node@" + "b"*40 + "\n"), 1, "HR-26"),
            ("pip install without ==", yml(run="pip install requests\n          " + RUN_OK), 1, "HR-26"),
            ("pip -r without hashes", yml(run="pip install -r requirements.txt\n          " + RUN_OK), 1, "HR-26"),
            ("pip -c constraints unverifiable", yml(run="pip install -c constraints.txt\n          " + RUN_OK), 1, "HR-26"),
            ("finding in non-required job", yml(
                tail="  other:\n    steps:\n      - uses: actions/checkout@v4\n"), 1, "HR-26"),
        ])

    def test_negatives(self):
        self.table([
            ("pinned sha clean", yml(), 0, None),
            ("setup-python exact", yml(
                tail="      - uses: actions/setup-python@" + "a"*40 +
                     "\n        with:\n          python-version: '3.12.7'\n"), 0, None),
            ("pip install with ==", yml(run="pip install requests==2.31.0\n          " + RUN_OK), 0, None),
            ("pip install -r with hashes stays clean", yml(
                run="pip install -r requirements.txt --require-hashes\n          " + RUN_OK), 0, None),
            ("local action needs no sha", yml(uses="uses: ./.github/actions/local"), 0, None),
            ("docker mutable tag", yml(uses="uses: docker://alpine:3.20"), 1, "HR-26"),
            ("docker digest clean", yml(uses="uses: docker://alpine@sha256:" + "c"*64), 0, None),
            ("degenerate job without steps", yml(run=None, uses=None,
                tail="  other:\n    steps:\n      - " + PINNED + "\n"), 0, None),
        ])


class HR3FailLoud(GateCase):
    """HR-3: masking patterns in required jobs fail loudly; scope is required jobs only."""
    def test_findings(self):
        self.table([
            ("continue-on-error step", yml().replace("      - " + PINNED,
             "      - continue-on-error: true\n      - " + PINNED), 1, "HR-3"),
            ("continue-on-error job", yml().replace("  build:", "  build:\n    continue-on-error: true"), 1, "HR-3"),
            ("continue-on-error workflow", yml().replace("on:", "continue-on-error: true\non:"), 1, "HR-3"),
            ("continue-on-error conditional", yml().replace("  build:",
             "  build:\n    continue-on-error: ${{ failure() }}"), 1, "HR-3"),
            ("or true", yml(run="true || true\n          " + RUN_OK), 1, "HR-3"),
            ("or colon", yml(run="set -u || :\n          " + RUN_OK), 1, "HR-3"),
            ("set +e", yml(run="set -u\n          set +e\n          " + RUN_OK), 1, "HR-3"),
            ("bare exit 0", yml(run="exit 0"), 1, "HR-3"),
        ])

    def test_negatives(self):
        self.table([
            ("clean required job", yml(), 0, None),
            ("explicit false stays clean", yml().replace("      - " + PINNED,
             "      - continue-on-error: false\n      - " + PINNED), 0, None),
            ("masking outside required jobs", yml(
                tail="  other:\n    steps:\n      - " + PINNED + "\n        run: echo ok || true\n"), 0, None),
            ("degenerate required job without steps", yml(run=None, uses=None), 0, None),
        ])


class HR30PrEventOnly(GateCase):
    """HR-30: required check names publish only from PR-evaluating events."""
    def test_findings(self):
        self.table([
            ("schedule with static name", yml(name="build",
                on="on:\n  pull_request:\n  schedule:\n    - cron: '0 3 * * *'\n"), 1, "HR-30"),
            ("dispatch with static name", yml(name="build",
                on="on:\n  pull_request:\n  workflow_dispatch:\n"), 1, "HR-30"),
            ("no pull_request trigger", yml(on="on:\n  push:\n    branches: [main]\n"), 1, "HR-30"),
            ("push to non-base branch", yml(name="build",
                on="on:\n  pull_request:\n  push:\n    branches: [develop]\n"), 1, "HR-30"),
            ("push without branch filter", yml(name="build", on="on:\n  pull_request:\n  push:\n"), 1, "HR-30"),
        ])

    def test_negatives(self):
        self.table([
            ("event-qualified name", yml(
                on="on:\n  pull_request:\n  schedule:\n    - cron: '0 3 * * *'\n  workflow_dispatch:\n"), 0, None),
            ("job if excludes non-PR events", yml(name="build",
                job="  build:\n    if: github.event_name == 'pull_request' || github.event_name == 'push'",
                on="on:\n  pull_request:\n  push:\n    branches: [main]\n  workflow_dispatch:\n"), 0, None),
            ("push base only", yml(name="build", on="on:\n  pull_request:\n  push:\n    branches: [main]\n"), 0, None),
            ("non-required static job beside required one", yml(
                on="on:\n  pull_request:\n  schedule:\n    - cron: '0 3 * * *'\n",
                tail="  other:\n    steps:\n      - " + PINNED + "\n        run: echo ok\n"), 0, None),
            ("degenerate on as scalar list", yml(on="on: [pull_request]"), 0, None),
        ])


class Cross(GateCase):
    def test_required_job_absent(self):
        code, out = self.run_gate("on:\n  pull_request:\n\njobs:\n  other:\n    steps:\n      - " + PINNED + "\n")
        self.assertEqual(code, 1)
        self.assertIn("build", out)

    def test_findings_aggregate(self):
        code, out = self.run_gate(yml(uses="uses: actions/checkout@v4"))
        self.assertEqual(code, 1)
        self.assertEqual(out.count("HR-26"), 1)

    def test_reader_helpers(self):
        lines = gate.read_lines("a: 1  # comment\n# full comment\n\nb: 'x'\n")
        self.assertEqual([l.text for l in lines], ["a: 1", "b: 'x'"])
        self.assertEqual(gate.parse_scalar("'3.12.7'"), "3.12.7")
        self.assertTrue(gate.parse_scalar("true") is True)

    def test_reader_block_scalar_and_flow_seq(self):
        doc = gate.parse_workflow(yml(), "wf.yml")
        self.assertEqual(doc["jobs"]["build"]["steps"][0]["run"], "set -u\nexit \"$fail\"\n")
        doc2 = gate.parse_workflow(yml(on="on: [pull_request]"), "wf.yml")
        self.assertEqual(doc2["jobs"]["build"]["steps"][0]["uses"],
                         "actions/checkout@11bd71901bbe5b1630ceea73d27597364c9af683")


if __name__ == "__main__":
    unittest.main(verbosity=2)
