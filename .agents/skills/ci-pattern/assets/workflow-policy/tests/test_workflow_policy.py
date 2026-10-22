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
    s = f"{on}\nconcurrency:\n  group: wf-${{{{ github.ref }}}}\njobs:\n{head}{job}\n    name: {name}\n    timeout-minutes: 5\n    steps:\n"
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
        self.table([(item, t, 2, None) for item, t, _ in self.CASES])

    def test_invalid_policy(self):
        for label, pol in [("no required_jobs", {"events": {}}),
                           ("empty required_jobs", {"required_jobs": [], "events": {}})]:
            with self.subTest(case=label):
                code, out = self.run_gate(yml(), policy=pol)
                self.assertEqual(code, 2)
                self.assertIn("required_jobs", out)

    def test_multiline_plain_run_names_the_line_and_the_fix(self):
        # #787 S3: `run: cmd \` en escalar plano. YAML pliega las lineas y el `\` queda literal,
        # asi que los argumentos nunca llegan (APAP_WEB runner-guard: 35 pushes rojos sin evaluar).
        text = yml(run=None) + (
            "        run: python3 check.py \\\n"
            "               --contract c.json \\\n"
            "               --event push\n")
        code, out = self.run_gate(text)
        self.assertEqual(code, 2, out)
        self.assertIn("line 12", out)
        self.assertIn("'run'", out)
        self.assertIn("block scalar", out)
        self.assertNotIn("stray content", out)

    def test_block_run_with_continuation_stays_valid(self):
        text = yml(run="python3 check.py \\\n            --contract c.json")
        self.assertEqual(self.run_gate(text)[0], 0, self.run_gate(text)[1])

    def test_nonexistent_workflow(self):
        self.assertEqual(self.run_gate(yml(), extra=("--workflow", "/no/such.yml"))[0], 2)

    def test_missing_required_argument(self):
        out = subprocess.run([sys.executable, str(GATE)], capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)


class TestHR26PinnedToolchain(GateCase):
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
                tail="  other:\n    timeout-minutes: 5\n    steps:\n      - uses: actions/checkout@v4\n"), 1, "HR-26"),
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
                tail="  other:\n    timeout-minutes: 5\n    steps:\n      - " + PINNED + "\n"), 0, None),
        ])


class TestHR3FailLoud(GateCase):
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
                tail="  other:\n    timeout-minutes: 5\n    steps:\n      - " + PINNED + "\n        run: echo ok || true\n"), 0, None),
            ("degenerate required job without steps", yml(run=None, uses=None), 0, None),
        ])


class TestHR30PrEventOnly(GateCase):
    """HR-30: required check names publish only from PR-evaluating events."""
    def test_findings(self):
        self.table([
            ("schedule with static name", yml(name="build",
                on="on:\n  pull_request:\n  schedule:\n    - cron: '0 3 * * *'\n"), 1, "HR-30"),
            ("dispatch with static name", yml(name="build",
                on="on:\n  pull_request:\n  workflow_dispatch:\n"), 1, "HR-30"),
            ("no pull_request trigger", yml(on="on:\n  push:\n    branches: [main]\n"), 1, "HR-30"),
            ("mixed PR and dispatch job (mutation shape)", yml(name="build",
                job="  build:\n    if: github.event_name == 'pull_request' || github.event_name == 'workflow_dispatch'",
                on="on:\n  pull_request:\n  workflow_dispatch:\n"), 1, "HR-30"),
            ("push without branch filter", yml(name="build", on="on:\n  pull_request:\n  push:\n"), 1, "HR-30"),
        ])

    def test_negatives(self):
        self.table([
            ("event-qualified name", yml(
                on="on:\n  pull_request:\n  schedule:\n    - cron: '0 3 * * *'\n  workflow_dispatch:\n"), 0, None),
            ("job if excludes non-PR events", yml(name="build",
                job="  build:\n    if: github.event_name == 'pull_request' || github.event_name == 'push'",
                on="on:\n  pull_request:\n  push:\n    branches: [main]\n  workflow_dispatch:\n"), 0, None),
            ("job if excludes PR, schedule and dispatch only (security-deep shape)", yml(name="build",
                job="  build:\n    if: github.event_name == 'workflow_dispatch' || github.event_name == 'schedule'",
                on="on:\n  pull_request:\n  schedule:\n    - cron: '0 3 * * *'\n  workflow_dispatch:\n"), 0, None),
            ("push base only", yml(name="build", on="on:\n  pull_request:\n  push:\n    branches: [main]\n"), 0, None),
            ("non-required static job beside required one", yml(
                on="on:\n  pull_request:\n  schedule:\n    - cron: '0 3 * * *'\n",
                tail="  other:\n    timeout-minutes: 5\n    steps:\n      - " + PINNED + "\n        run: echo ok\n"), 0, None),
            ("degenerate on as scalar list", yml(on="on: [pull_request]"), 0, None),
        ])


SETUP_PY = ("      - uses: actions/setup-python@" + "a" * 40 +
            "\n        with:\n          python-version-file: {}\n")


class RepoCase(GateCase):
    """Gate run inside a temp repo layout so files named by a workflow resolve."""
    def run_repo(self, text, files=None, flag=True, rel="wf.yml", contract=None):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            for name, content in (files or {}).items():
                target = root / name
                target.parent.mkdir(parents=True, exist_ok=True)
                if content is None:
                    target.mkdir()
                else:
                    target.write_text(content, encoding="utf-8")
            w, p = root / rel, root / "policy.json"
            w.parent.mkdir(parents=True, exist_ok=True)
            w.write_text(text, encoding="utf-8")
            p.write_text(json.dumps(POLICY), encoding="utf-8")
            extra = ["--repo-root", str(root)] if flag else []
            if contract is not None:
                c = root / "host-contract.json"
                c.write_text(contract if isinstance(contract, str) else json.dumps(contract),
                             encoding="utf-8")
                extra += ["--host-contract", str(c)]
            out = subprocess.run([sys.executable, str(GATE), "--workflow", str(w),
                                  "--required-jobs", str(p), "--base-branch", "main", *extra],
                                 capture_output=True, text=True)
        return out.returncode, out.stdout + out.stderr


class TestHR26PythonVersionFile(RepoCase):
    """HR-26: python-version-file counts as exact when the file holds exactly X.Y.Z."""
    def wf(self, ref=".python-version", inline=None):
        extra = f"          python-version: '{inline}'\n" if inline else ""
        return yml(tail=SETUP_PY.format(ref) + extra)

    def test_exact_file_is_clean(self):
        for label, content in (("plain", "3.12.14"), ("trailing newline", "3.12.14\n")):
            with self.subTest(case=label):
                code, out = self.run_repo(self.wf(), {".python-version": content})
                self.assertEqual(code, 0, out)

    def test_repo_root_defaults_to_the_workflow_checkout(self):
        code, out = self.run_repo(self.wf(), {".python-version": "3.12.14"}, flag=False,
                                  rel=".github/workflows/ci.yml")
        self.assertEqual(code, 0, out)

    def test_fails_closed(self):
        cases = (
            ("file missing", self.wf(), {}),
            ("minor only", self.wf(), {".python-version": "3.12"}),
            ("range", self.wf(), {".python-version": ">=3.12.0"}),
            ("two lines", self.wf(), {".python-version": "3.12.14\n3.11.0\n"}),
            ("empty file", self.wf(), {".python-version": ""}),
            ("directory", self.wf(), {".python-version": None}),
            ("escapes the repo", self.wf("../outside"), {}),
            ("absolute path", self.wf("/etc/hostname"), {}),
            ("expression", self.wf("${{ matrix.file }}"), {".python-version": "3.12.14"}),
            ("inline wins and is not exact", self.wf(inline="3.12"), {".python-version": "3.12.14"}),
        )
        for label, text, files in cases:
            with self.subTest(case=label):
                code, out = self.run_repo(text, files)
                self.assertEqual(code, 1, out)
                self.assertIn("HR-26", out)


def push_wf(push, name="build", job="  build:", extra_on=""):
    return yml(name=name, job=job, on="on:\n  pull_request:\n" + extra_on + push)


class TestHR30PushFilters(GateCase):
    """HR-30: a push that never reaches the base publishes no PR check."""
    def test_negatives(self):
        self.table([
            ("push staging only", push_wf("  push:\n    branches: [staging]\n"), 0, None),
            ("push tags only", push_wf("  push:\n    tags: ['v*']\n"), 0, None),
            ("push staging and tags", push_wf("  push:\n    branches: [staging]\n    tags: ['v*']\n"), 0, None),
            ("push branches-ignore the base", push_wf("  push:\n    branches-ignore: [main]\n"), 0, None),
            ("tags-ignore only runs on tag pushes", push_wf("  push:\n    tags-ignore: ['v*']\n"), 0, None),
            ("push base only plus tags stays clean", push_wf(
                "  push:\n    branches: [main]\n    tags: ['v*']\n"), 0, None),
        ])

    def test_findings(self):
        self.table([
            ("base and another branch", push_wf("  push:\n    branches: [main, staging]\n"), 1, "HR-30"),
            ("glob reaching the base", push_wf("  push:\n    branches: ['**']\n"), 1, "HR-30"),
            ("branches-ignore not covering the base", push_wf(
                "  push:\n    branches-ignore: ['docs/**']\n"), 1, "HR-30"),
            ("negation is not provable", push_wf("  push:\n    branches: ['!main']\n"), 1, "HR-30"),
            ("branches and branches-ignore together", push_wf(
                "  push:\n    branches: [staging]\n    branches-ignore: [main]\n"), 1, "HR-30"),
            ("still no filter", push_wf("  push:\n"), 1, "HR-30"),
        ])


SD_ON = ("on:\n  pull_request:\n  schedule:\n    - cron: '0 3 * * *'\n  workflow_dispatch:\n")


def if_wf(cond, on=SD_ON):
    return yml(name="build", job="  build:\n    if: " + cond, on=on)


class TestHR30JobConditions(GateCase):
    """HR-30: a job provably skipped on pull_request publishes no PR check name."""
    def test_negatives(self):
        folded = ("  build:\n    if: >-\n      !cancelled() &&\n      (\n"
                  "        github.event_name == 'workflow_dispatch' ||\n"
                  "        github.event_name == 'schedule'\n      ) &&\n"
                  "      (needs.pr-gate.result == 'success' || needs.pr-gate.result == 'skipped')")
        self.table([
            ("not pull_request", if_wf("github.event_name != 'pull_request'"), 0, None),
            ("negated group", if_wf("\"!(github.event_name == 'pull_request')\""), 0, None),
            ("not pull_request and an unknown atom", if_wf(
                "github.event_name != 'pull_request' && needs.pr-gate.result == 'success'"), 0, None),
            ("security-deep shape, folded and compound", yml(name="build", job=folded, on=SD_ON), 0, None),
            ("allowlist without pull_request and unknown conjunct", if_wf(
                "always() && (github.event_name == 'schedule' || github.event_name == 'workflow_dispatch')"),
             0, None),
        ])

    def test_findings(self):
        self.table([
            ("typecheck shape runs on PR when not docs-only", if_wf(
                "github.event_name != 'pull_request' || needs.pr-gate.outputs.docs_changed != 'true'"),
             1, "HR-30"),
            ("unknown disjunct keeps PR reachable", if_wf(
                "github.event_name == 'schedule' || needs.pr-gate.result == 'success'"), 1, "HR-30"),
            ("not schedule still runs on PR and dispatch", if_wf("github.event_name != 'schedule'"), 1, "HR-30"),
            ("unbalanced group stays closed", if_wf("(github.event_name != 'pull_request'"), 1, "HR-30"),
            ("expression wrapper stays closed", if_wf("${{ github.event_name != 'pull_request' }}"), 1, "HR-30"),
            ("unknown only", if_wf("always()"), 1, "HR-30"),
        ])


def contract_of(*names):
    return {"contract_version": 1,
            "required_checks": [{"name": n, "class": "documented-only"} for n in names]}


AGG = ("  required:\n    name: required\n    timeout-minutes: 5\n    steps:\n      - " + PINNED + "\n        run: echo ok\n")


class TestHR30DeclaredRequiredChecks(RepoCase):
    """HR-30 schedule/dispatch applies only to the check names the host contract requires."""
    def wf(self, name="build", job="  build:", tail="", on=SD_ON):
        return yml(name=name, job=job, on=on, tail=tail)

    def check(self, text, contract, code, needle=None):
        got, out = self.run_repo(text, contract=contract)
        self.assertEqual(got, code, out)
        if needle:
            self.assertIn(needle, out)

    def test_undeclared_job_publishing_on_schedule_and_dispatch_is_clean(self):
        qualified_agg = AGG.replace("name: required", "name: required (${{ github.event_name }})")
        self.check(self.wf(tail=qualified_agg), contract_of("required (pull_request)"), 0)

    def test_declared_job_still_findings(self):
        self.check(self.wf(), contract_of("build"), 1, "HR-30")
        self.check(self.wf(tail=AGG), contract_of("required"), 1, "'required'")

    def test_aggregator_outside_the_policy_list_is_checked_when_declared(self):
        got, out = self.run_repo(self.wf(tail=AGG), contract=contract_of("required"))
        self.assertEqual(got, 1, out)
        self.assertIn("job 'required'", out)
        self.assertNotIn("job 'build'", out)

    def test_declared_names_that_satisfy_the_rule(self):
        qualified = self.wf(name="build (${{ github.event_name }})")
        self.check(qualified, contract_of("build (pull_request)"), 0)
        excluded = self.wf(job="  build:\n    if: github.event_name != 'pull_request'")
        self.check(excluded, contract_of("build"), 0)

    def test_empty_declaration_exempts_the_policy_jobs(self):
        self.check(self.wf(), contract_of(), 0)
        self.check(self.wf(), {"contract_version": 1}, 0)

    def test_declared_name_no_job_publishes_is_a_finding(self):
        self.check(self.wf(), contract_of("build", "lint"), 1, "requires check 'lint'")

    def test_without_the_option_every_policy_job_is_checked(self):
        got, out = self.run_repo(self.wf())
        self.assertEqual(got, 1, out)
        self.assertIn("HR-30", out)

    def test_invalid_contract_exits_2(self):
        for label, bad in (("not json", "{"), ("not an object", "[]"),
                           ("required_checks not a list", {"required_checks": {}}),
                           ("entry without name", {"required_checks": [{"class": "x"}]}),
                           ("empty name", {"required_checks": [{"name": ""}]})):
            with self.subTest(case=label):
                self.check(self.wf(), bad, 2)


SCHED_PINNED = "actions/checkout@11bd71901bbe5b1630ceea73d27597364c9af683"


GOOD_BUILD = """  build:
    name: build (${{ github.event_name }})
    timeout-minutes: 5
    concurrency: j
    steps:
      - uses: """ + SCHED_PINNED + """
        run: |
          set -u
          exit "$fail"
"""

SECURITY_TAIL = """  security:
    name: security
    timeout-minutes: 5
    concurrency: j
    steps:
      - uses: """ + SCHED_PINNED + """
      - run: pip-audit
"""

AUDIT_TAIL_GOOD = """  audit:
    name: audit
    timeout-minutes: 5
    concurrency: j
    steps:
      - uses: """ + SCHED_PINNED + """
      - run: pip-audit
      - run: gh issue create --title audit --body x
"""

AUDIT_TAIL_NO_ALERT = """  audit:
    name: audit
    timeout-minutes: 5
    concurrency: j
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


class TestHR39StateDependentGates(unittest.TestCase):
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


class TestDuplicateKeys(GateCase):
    """#806 C1: GitHub rejects a duplicate mapping key with `startup_failure`, so the check
    never appears in the PR rollup. The reader keeps the last value (PyYAML does too), so it
    has to refuse the file instead of guessing."""

    def test_duplicate_key_in_a_step_is_refused_with_both_lines(self):
        text = yml(uses=PINNED + "\n        with:\n          fetch-depth: 0\n        with:\n"
                   "          fetch-depth: 1")
        code, out = self.run_gate(text)
        self.assertEqual(code, 2, out)
        self.assertIn("duplicate key 'with'", out)
        self.assertIn("line 14", out)
        self.assertIn("line 12", out)

    def test_duplicate_job_id_is_refused(self):
        text = yml(tail="  build:\n    timeout-minutes: 5\n    steps:\n      - " + PINNED + "\n")
        code, out = self.run_gate(text)
        self.assertEqual(code, 2, out)
        self.assertIn("duplicate key 'build'", out)

    def test_same_key_in_sibling_mappings_is_clean(self):
        text = yml(uses=PINNED + "\n        with:\n          fetch-depth: 0\n      - "
                   + PINNED + "\n        with:\n          fetch-depth: 1")
        code, out = self.run_gate(text)
        self.assertEqual(code, 0, out)


class TestJobTimeout(GateCase):
    """#806 C1: GitHub's default job budget is 360 minutes; a wedged runner then holds the
    queue for six hours. Every job states its own, required or not."""
    NO_TIMEOUT = ("on:\n  pull_request:\nconcurrency:\n  group: wf\n\njobs:\n  build:\n    name: build (${{ github.event_name }})\n"
                  "    steps:\n      - " + PINNED + "\n")

    def test_job_without_timeout_is_a_finding(self):
        code, out = self.run_gate(self.NO_TIMEOUT)
        self.assertEqual(code, 1, out)
        self.assertIn("job 'build'", out)
        self.assertIn("timeout-minutes", out)

    def test_job_that_is_not_required_is_also_a_finding(self):
        text = self.NO_TIMEOUT + "  extra:\n    timeout-minutes: 5\n    steps:\n      - " + PINNED + "\n"
        text = text.replace("  build:\n", "  build:\n    timeout-minutes: 5\n", 1)
        text += "  other:\n    steps:\n      - " + PINNED + "\n"
        code, out = self.run_gate(text)
        self.assertEqual(code, 1, out)
        self.assertIn("job 'other'", out)
        self.assertNotIn("job 'extra'", out)

    def test_job_with_timeout_is_clean(self):
        self.assertEqual(self.run_gate(yml())[0], 0)

    def test_reusable_workflow_call_is_exempt(self):
        # The schema of a job that calls a reusable workflow has no `timeout-minutes`.
        text = yml(tail="  call:\n    uses: ./.github/workflows/other.yml\n")
        code, out = self.run_gate(text)
        self.assertEqual(code, 0, out)


NO_CONCURRENCY = ("on:\n  pull_request:\n\njobs:\n  build:\n    name: build (${{ github.event_name }})\n"
                  "    timeout-minutes: 5\n    steps:\n      - " + PINNED + "\n")


class TestConcurrency(GateCase):
    """#806 C2: without a concurrency group a second push runs alongside the first and both
    compete for the same runner. A workflow-level group covers every job."""

    def test_no_group_anywhere_is_a_finding(self):
        code, out = self.run_gate(NO_CONCURRENCY)
        self.assertEqual(code, 1, out)
        self.assertIn("job 'build'", out)
        self.assertIn("concurrency", out)

    def test_workflow_level_group_covers_every_job(self):
        text = "concurrency:\n  group: wf\n" + NO_CONCURRENCY
        self.assertEqual(self.run_gate(text)[0], 0)

    def test_job_level_group_covers_that_job(self):
        text = NO_CONCURRENCY.replace("    timeout-minutes: 5\n",
                                      "    timeout-minutes: 5\n    concurrency:\n      group: build\n")
        self.assertEqual(self.run_gate(text)[0], 0)

    def test_string_form_counts_as_a_group(self):
        text = NO_CONCURRENCY.replace("    timeout-minutes: 5\n",
                                      "    timeout-minutes: 5\n    concurrency: build\n")
        self.assertEqual(self.run_gate(text)[0], 0)

    def test_mapping_without_group_is_a_finding(self):
        text = "concurrency:\n  cancel-in-progress: true\n" + NO_CONCURRENCY
        code, out = self.run_gate(text)
        self.assertEqual(code, 1, out)
        self.assertIn("concurrency", out)


class TestServicePorts(GateCase):
    """#806 C2: `5432:5432` reserves a host port on the runner; with one runner nothing collides
    and the defect stays invisible, with two the jobs fail to bind or share a database."""

    def with_ports(self, *ports):
        listing = "".join(f"          - {p}\n" for p in ports)
        return yml(tail="    services:\n      db:\n        image: postgres@sha256:" + "a" * 64
                   + "\n        ports:\n" + listing)

    def test_fixed_host_port_is_a_finding(self):
        code, out = self.run_gate(self.with_ports("5432:5432"))
        self.assertEqual(code, 1, out)
        self.assertIn("service 'db'", out)
        self.assertIn("5432:5432", out)

    def test_host_ip_binding_is_also_fixed(self):
        self.assertEqual(self.run_gate(self.with_ports("127.0.0.1:5432:5432"))[0], 1)

    def test_container_port_alone_is_clean(self):
        code, out = self.run_gate(self.with_ports("5432"))
        self.assertEqual(code, 0, out)

    def test_quoted_and_flow_forms_are_read(self):
        self.assertEqual(self.run_gate(self.with_ports('"5432:5432"'))[0], 1)
        flow = yml(tail="    services:\n      db:\n        image: x\n        ports: [5432:5432]\n")
        self.assertEqual(self.run_gate(flow)[0], 1)


class TestDockerPreflight(GateCase):
    """#806 C2: `docker run` against a wedged daemon blocks instead of failing. The guard has to
    be `docker info` under `timeout`, earlier in the job or earlier in the same step."""

    def job_with(self, *runs):
        body = "".join(f"      - run: |\n          {r}\n" for r in runs)
        return ("on:\n  pull_request:\nconcurrency:\n  group: wf\n\njobs:\n  build:\n"
                "    name: build (${{ github.event_name }})\n    timeout-minutes: 5\n    steps:\n"
                "      - " + PINNED + "\n" + body)

    def test_docker_run_without_guard_is_a_finding(self):
        code, out = self.run_gate(self.job_with("docker run --rm img"))
        self.assertEqual(code, 1, out)
        self.assertIn("docker run", out)
        self.assertIn("job 'build'", out)

    def test_guard_in_an_earlier_step_is_clean(self):
        text = self.job_with("timeout 30 docker info > /dev/null", "docker run --rm img")
        self.assertEqual(self.run_gate(text)[0], 0)

    def test_guard_earlier_in_the_same_step_is_clean(self):
        text = self.job_with("timeout 30 docker info > /dev/null\n          docker run --rm img")
        self.assertEqual(self.run_gate(text)[0], 0)

    def test_bare_docker_info_hangs_like_no_guard(self):
        code, out = self.run_gate(self.job_with("docker info", "docker run --rm img"))
        self.assertEqual(code, 1, out)
        self.assertIn("timeout", out)

    def test_guard_after_the_run_does_not_count(self):
        text = self.job_with("docker run --rm img\n          timeout 30 docker info")
        self.assertEqual(self.run_gate(text)[0], 1)

    def test_guard_only_in_a_comment_does_not_count(self):
        text = self.job_with("# timeout 30 docker info\n          docker run --rm img")
        self.assertEqual(self.run_gate(text)[0], 1)

    def test_job_without_docker_run_is_clean(self):
        self.assertEqual(self.run_gate(self.job_with("echo hi"))[0], 0)


class TestCross(GateCase):
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
        self.assertEqual([item.text for item in lines], ["a: 1", "b: 'x'"])
        self.assertEqual(gate.parse_scalar("'3.12.7'"), "3.12.7")
        self.assertTrue(gate.parse_scalar("true") is True)

    def test_reader_block_scalar_and_flow_seq(self):
        doc = gate.parse_workflow(yml(), "wf.yml")
        self.assertEqual(doc["jobs"]["build"]["steps"][0]["run"], "set -u\nexit \"$fail\"\n")
        doc2 = gate.parse_workflow(yml(on="on: [pull_request]"), "wf.yml")
        self.assertEqual(doc2["jobs"]["build"]["steps"][0]["uses"],
                         "actions/checkout@11bd71901bbe5b1630ceea73d27597364c9af683")


class TestPullRequestTypesListStyle(GateCase):
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
        "    timeout-minutes: 5\n"
        "    concurrency: j\n"
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


class TestWorkflowCallEmptyInputsDict(GateCase):
    """RED test para el bug del parser: `on.workflow_call.inputs: {}`
    (mapa de flujo vacío) lanza 'unsupported YAML construct: {}' (rc=2)
    porque `parse_scalar` rechaza cualquier escalar que empiece por `&`,
    `!` o `{`. GitHub Actions admite `inputs: {}` y `secrets: {}` como
    mapeos vacíos en un reusable workflow; el canon debe aceptarlos
    como conformantes. Hoy cadete (#533, #806 K6) lo sufre y esquiva el
    problema quitando el workflow de los args del gate; este test
    reproduce el error de raíz."""

    WORKFLOW = (
        "on:\n"
        "  workflow_call:\n"
        "    inputs: {}\n"
        "\n"
        "concurrency:\n"
        "  group: empty-${{ github.ref }}\n"
        "\n"
        "jobs:\n"
        "  call:\n"
        "    name: call (${{ github.event_name }})\n"
        "    runs-on: ubuntu-latest\n"
        "    timeout-minutes: 5\n"
        "    steps:\n"
        "      - run: |\n"
        "          echo ok\n"
    )

    def test_workflow_call_empty_inputs_dict_parses_clean(self):
        # Hoy el gate sale 2 con 'unsupported YAML construct: '{}''.
        # Tras el fix, debe parsear como una pieza conforme: rc distinto
        # de 2 y sin el literal 'unsupported YAML construct: {}'. La
        # verificación del veredicto concreto (rc=0 con un hallazgo o
        # rc=1 con un HR concreto) pertenece a tests de reglas; aquí
        # verificamos solo que el parser NO rechaza `inputs: {}`.
        code, out = self.run_gate(self.WORKFLOW)
        self.assertNotEqual(
            code, 2,
            f"el parser rechazó `on.workflow_call.inputs: {{}}` con rc=2:\n{out}",
        )
    def test_inline_map_with_content_fails_closed(self):
        # Complemento del test anterior: un mapa inline CON contenido
        # (p. ej. `inputs: {a: b}`) NO se soporta — el reader solo
        # acepta `{}`. La forma correcta es bloque indentado
        # (`inputs:\n  a: b`) o `inputs: {}` si está vacío. El reader
        # debe fallar cerrado con rc=2 y un mensaje que nombre el
        # mapa rechazado y la alternativa soportada.
        wf = (
            "on:\n"
            "  workflow_call:\n"
            "    inputs: {a: b}\n"
            "\n"
            "concurrency:\n"
            "  group: x-${{ github.ref }}\n"
            "\n"
            "jobs:\n"
            "  call:\n"
            "    name: call (${{ github.event_name }})\n"
            "    runs-on: ubuntu-latest\n"
            "    timeout-minutes: 5\n"
            "    steps:\n"
            "      - run: |\n"
            "          echo ok\n"
        )
        code, out = self.run_gate(wf)
        self.assertEqual(
            code, 2,
            f"el parser NO falló cerrado ante `{{a: b}}`:\n{out}",
        )
        self.assertIn(
            "inline map", out,
            f"el mensaje de error no nombra el constructo rechazado:\n{out}",
        )
        self.assertIn(
            "{a: b}", out,
            f"el mensaje no reproduce el mapa rechazado para diagnóstico:\n{out}",
        )

if __name__ == "__main__":
    unittest.main(verbosity=2)
