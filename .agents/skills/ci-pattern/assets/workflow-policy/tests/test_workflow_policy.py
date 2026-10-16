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

POLICY = {"required_jobs": ["build"], "events": {"pull_request": {"accepted_skips": []}},
           "state_dependent_gates": []}
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


SCHED_PINNED = "actions/checkout@11bd71901bbe5b1630ceea73d27597364c9af683"


GOOD_BUILD = """  build:
    name: build (${{ github.event_name }})
    steps:
      - uses: """ + SCHED_PINNED + """
        run: |
          set -u
          exit "$fail"
"""

SECURITY_TAIL = """  security:
    name: security
    steps:
      - uses: """ + SCHED_PINNED + """
      - run: pip-audit
"""

AUDIT_TAIL_GOOD = """  audit:
    name: audit
    steps:
      - uses: """ + SCHED_PINNED + """
      - run: pip-audit
      - run: gh issue create --title audit --body x
"""

AUDIT_TAIL_NO_ALERT = """  audit:
    name: audit
    steps:
      - uses: """ + SCHED_PINNED + """
      - run: pip-audit
"""


def pr_wf(paths_flow="[uv.lock, requirements*.txt]"):
    if paths_flow:
        on = "on:\n  pull_request:\n    paths: " + paths_flow + "\n"
    else:
        on = "on:\n  pull_request:\n"
    return on + "jobs:\n" + GOOD_BUILD + SECURITY_TAIL


def sched_wf(tail=AUDIT_TAIL_GOOD,
             on_block="on:\n  schedule:\n    - cron: '0 3 * * *'\n  push:\n    branches: [main]\n"):
    return on_block + "jobs:\n" + tail


GOOD_STATE_ENTRY = {
    "name": "dependency-audit",
    "command": "pip-audit",
    "pr_workflow": "wf_pr.yml",
    "pr_job": "security",
    "manifests": ["uv.lock", "requirements*.txt"],
    "scheduled_workflow": "wf_sched.yml",
    "alert_command": "gh issue create",
}

REQUIRED_POLICY = {"required_jobs": ["build"]}


class HR39StateDependentGates(unittest.TestCase):
    """HR-39: un gate dependiente de estado externo va programado y solo
    bloquea PRs que tocan los manifiestos declarados."""

    def run_pair(self, pr_text, sched_text, policy=None):
        pol = dict(REQUIRED_POLICY)
        if policy is not None:
            pol.update(policy)
        with tempfile.TemporaryDirectory() as d:
            d_path = Path(d)
            p = d_path / "policy.json"
            pr_path = d_path / "wf_pr.yml"
            sched_path = d_path / "wf_sched.yml"
            p.write_text(json.dumps(pol), encoding="utf-8")
            pr_path.write_text(pr_text, encoding="utf-8")
            sched_path.write_text(sched_text, encoding="utf-8")
            cmd = [sys.executable, str(GATE),
                   "--workflow", "wf_pr.yml",
                   "--workflow", "wf_sched.yml",
                   "--required-jobs", "policy.json", "--base-branch", "main"]
            out = subprocess.run(cmd, capture_output=True, text=True, cwd=str(d_path))
        return out.returncode, out.stdout, out.stderr

    def run_pair_no_sched(self, pr_text, policy):
        pol = dict(REQUIRED_POLICY)
        pol.update(policy)
        with tempfile.TemporaryDirectory() as d:
            d_path = Path(d)
            p = d_path / "policy.json"
            pr_path = d_path / "wf_pr.yml"
            p.write_text(json.dumps(pol), encoding="utf-8")
            pr_path.write_text(pr_text, encoding="utf-8")
            cmd = [sys.executable, str(GATE),
                   "--workflow", "wf_pr.yml",
                   "--required-jobs", "policy.json", "--base-branch", "main"]
            out = subprocess.run(cmd, capture_output=True, text=True, cwd=str(d_path))
        return out.returncode, out.stdout, out.stderr

    def assert_clean_hr39(self, code, out, err):
        self.assertEqual(code, 0, (out, err))
        self.assertNotIn("HR-39", out)
        self.assertNotIn("HR-39", err)

    def test_clean_declaration_exits_zero(self):
        """(a) declaracion limpia: PR con paths cubriendo manifiestos y
        schedule con audit + alert: verde."""
        pol = {"state_dependent_gates": [GOOD_STATE_ENTRY]}
        code, out, err = self.run_pair(pr_wf(), sched_wf(), policy=pol)
        self.assert_clean_hr39(code, out, err)

    def test_pull_request_without_paths_is_finding(self):
        """(b) pull_request sin paths: el gate bloquea PRs que no tocan
        manifiestos. Finding HR-39."""
        pol = {"state_dependent_gates": [GOOD_STATE_ENTRY]}
        code, out, err = self.run_pair(pr_wf(paths_flow=""), sched_wf(), policy=pol)
        self.assertEqual(code, 1, (out, err))
        self.assertIn("HR-39", out)
        self.assertIn("paths", out)

    def test_manifest_not_covered_is_finding(self):
        """(c) un manifiesto declarado que ningun path pattern cubre: finding."""
        pol = {"state_dependent_gates": [GOOD_STATE_ENTRY]}
        code, out, err = self.run_pair(
            pr_wf(paths_flow="[pyproject.toml]"), sched_wf(), policy=pol)
        self.assertEqual(code, 1, (out, err))
        self.assertIn("HR-39", out)
        self.assertIn("uv.lock", out)
        self.assertIn("requirements", out)

    def test_no_scheduled_workflow_is_finding(self):
        """(d) el policy declara scheduled_workflow pero no aparece en inputs: finding."""
        pol = {"state_dependent_gates": [GOOD_STATE_ENTRY]}
        code, out, err = self.run_pair_no_sched(pr_wf(), policy=pol)
        self.assertEqual(code, 1, (out, err))
        self.assertIn("HR-39", out)
        self.assertIn("wf_sched.yml", out)

    def test_scheduled_without_alert_is_finding(self):
        """(e) el workflow programado no ejecuta el alert_command: finding."""
        pol = {"state_dependent_gates": [GOOD_STATE_ENTRY]}
        code, out, err = self.run_pair(
            pr_wf(), sched_wf(tail=AUDIT_TAIL_NO_ALERT), policy=pol)
        self.assertEqual(code, 1, (out, err))
        self.assertIn("HR-39", out)
        self.assertIn("gh issue create", out)

    def test_key_absent_exits_zero_with_notice(self):
        """(f) state_dependent_gates ausente: el gate imprime notice por
        stderr y sale 0 (fail-open; lo cierra por fuera el doctor)."""
        code, out, err = self.run_pair(pr_wf(), sched_wf())
        self.assertEqual(code, 0, (out, err))
        self.assertIn("HR-39", err)
        self.assertNotIn("HR-39", out)

    def test_empty_list_exits_zero(self):
        """Una lista vacia es declaracion valida: 'no tenemos gates
        dependientes de estado'. El propio repo del patron es dogfood de
        este caso."""
        pol = {"state_dependent_gates": []}
        code, out, err = self.run_pair(pr_wf(), sched_wf(), policy=pol)
        self.assertEqual(code, 0, (out, err))
        self.assertNotIn("HR-39", out)
        self.assertNotIn("HR-39", err)

    def test_malformed_entry_exits_two(self):
        """(g) entrada malformada (campo obligatorio ausente, manifests
        vacio, manifests no lista de cadenas): exit 2 sin nombrar HR-39."""
        bad_entries = [
            ({"name": "x"}, "missing-fields"),
            (dict(GOOD_STATE_ENTRY, manifests=[]), "manifests-empty"),
            (dict(GOOD_STATE_ENTRY, manifests="uv.lock"), "manifests-not-list"),
            (dict(GOOD_STATE_ENTRY, manifests=["uv.lock", 42]), "manifests-non-str"),
            (dict(GOOD_STATE_ENTRY, name=""), "name-empty"),
            (dict(GOOD_STATE_ENTRY, command=""), "command-empty"),
            (dict(GOOD_STATE_ENTRY, pr_workflow=""), "pr-workflow-empty"),
            (dict(GOOD_STATE_ENTRY, pr_job=""), "pr-job-empty"),
            (dict(GOOD_STATE_ENTRY, scheduled_workflow=""), "scheduled-empty"),
            (dict(GOOD_STATE_ENTRY, alert_command=""), "alert-empty"),
        ]
        for bad, label in bad_entries:
            with self.subTest(case=label):
                pol = {"state_dependent_gates": [bad]}
                code, out, err = self.run_pair(pr_wf(), sched_wf(), policy=pol)
                self.assertEqual(code, 2, (out, err))
                # El shape invalido es exit 2 por HR-3 y HR-18, no por HR-39.
                self.assertNotIn("HR-39", out)

    def test_top_level_must_be_list(self):
        pol = {"state_dependent_gates": {"name": "x"}}
        code, out, err = self.run_pair(pr_wf(), sched_wf(), policy=pol)
        self.assertEqual(code, 2, (out, err))
        self.assertIn("state_dependent_gates", err)


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


class PullRequestTypesListStyle(GateCase):
    """`pull_request.types:` admite dos formas equivalentes de YAML
    (lista en bloque y lista en flujo). El reader estructural del gate
    debe producir el mismo veredicto para ambas. La forma en bloque
    es la más común en workflows reales (#857 PR-B de BP) y el parser
    debe soportarla igual que la forma en flujo (que ya pasaba)."""

    BLOCK_TYPES = (
        "on:\n"
        "  pull_request:\n"
        "    types:\n"
        "      - opened\n"
        "      - synchronize\n"
        "      - reopened\n"
    )
    FLOW_TYPES = "on:\n  pull_request:\n    types: [opened, synchronize, reopened]\n"
    TAIL = (
        "  push:\n"
        "    branches: [main]\n"
        "jobs:\n"
        "  build:\n"
        "    name: build (${{ github.event_name }})\n"
        "    steps:\n"
        "      - " + PINNED + "\n"
        "        run: |\n"
        "          set -u\n"
        "          exit \"$fail\"\n"
    )

    def test_block_style_parses_clean(self):
        """La lista en bloque bajo `types:` no es 'unsupported construct':
        el reader la acepta y el gate sale 0 sobre el mismo cuerpo que
        sale 0 en la forma en flujo."""
        got, out = self.run_gate(self.BLOCK_TYPES + self.TAIL)
        self.assertEqual(got, 0, out)

    def test_block_and_flow_yield_same_verdict(self):
        """Dos workflows que difieren solo en la forma de `types:` deben
        producir el mismo veredicto del gate: 'mismo resultado'."""
        got_b, out_b = self.run_gate(self.BLOCK_TYPES + self.TAIL)
        got_f, out_f = self.run_gate(self.FLOW_TYPES + self.TAIL)
        self.assertEqual(
            got_b, got_f,
            f"block exit {got_b}: {out_b!r}\nflow exit {got_f}: {out_f!r}",
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
