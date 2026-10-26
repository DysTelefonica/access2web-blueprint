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
            ("bare name with extras is a registry requirement", yml(
                run="pip install \"app[dev]\"\n          " + RUN_OK), 1, "HR-26"),
            ("local path mixed with an unpinned registry name", yml(
                run="pip install ./app requests\n          " + RUN_OK), 1, "HR-26"),
            ("local path with a requirements file without hashes", yml(
                run="pip install ./app -r requirements.txt\n          " + RUN_OK), 1, "HR-26"),
            ("editable bare name is a registry requirement", yml(
                run="pip install -e app\n          " + RUN_OK), 1, "HR-26"),
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
            ("pip install of a local path with extras", yml(
                run="python -m pip install --quiet \"./app[dev]\"\n          " + RUN_OK), 0, None),
            ("pip install -e of a local path", yml(run="pip install -e ./app\n          " + RUN_OK), 0, None),
            ("pip install of the current directory", yml(run="pip install .\n          " + RUN_OK), 0, None),
            ("pip install of several local paths", yml(
                run="pip install --quiet ../x ./y /abs/z ~/w\n          " + RUN_OK), 0, None),
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


class TestHR30ReusableDeclaredNames(unittest.TestCase):
    """#1311: el nombre de check que publica un workflow reutilizable es compuesto
    (`<job que llama> / <job llamado>`), no el nombre del job solo.

    El gate lo reconoce solo cuando puede PROBARlo (el workflow llamado está entre
    los ficheros analizados); si no, falla cerrado: el nombre declarado sigue sin
    publicarse y el finding se mantiene."""

    PINNED = "actions/checkout@11bd71901bbe5b1630ceea73d27597364c9af683"

    #: El workflow que llama: su job `pr-size` llama al reusable con el mismo nombre.
    CALLER = (
        "name: pr-metadata\n"
        "on:\n"
        "  pull_request:\n"
        "concurrency:\n"
        "  group: wf-${{ github.ref }}\n"
        "jobs:\n"
        "  pr-size:\n"
        "    name: pr-size\n"
        "    uses: ./.github/workflows/pr-size.yml\n"
    )

    #: El workflow reutilizable: su job `pr-size` publica `pr-size / pr-size`.
    REUSABLE = (
        "name: pr-size\n"
        "on:\n"
        "  workflow_call:\n"
        "  pull_request:\n"
        "  schedule:\n"
        "    - cron: '0 3 * * *'\n"
        "concurrency:\n"
        "  group: reusable-${{ github.workflow }}\n"
        "jobs:\n"
        "  pr-size:\n"
        "    name: pr-size\n"
        "    timeout-minutes: 5\n"
        "    steps:\n"
        "      - uses: " + PINNED + "\n"
    )

    #: Control: el mismo nombre, publicado por un job propio del workflow.
    PLAIN = (
        "name: plain\n"
        "on:\n"
        "  pull_request:\n"
        "  schedule:\n"
        "    - cron: '0 3 * * *'\n"
        "concurrency:\n"
        "  group: wf-${{ github.ref }}\n"
        "jobs:\n"
        "  pr-size:\n"
        "    name: pr-size / pr-size\n"
        "    timeout-minutes: 5\n"
        "    steps:\n"
        "      - uses: " + PINNED + "\n"
    )

    #: La forma medida en access2web-blueprint: `security (pull_request) / gitleaks`.
    SECURITY_CALLER = (
        "name: security\n"
        "on:\n"
        "  pull_request:\n"
        "concurrency:\n"
        "  group: wf-${{ github.ref }}\n"
        "jobs:\n"
        "  security:\n"
        "    name: security (${{ github.event_name }})\n"
        "    uses: ./.github/workflows/security.yml\n"
    )

    SECURITY_REUSABLE = (
        "name: security\n"
        "on:\n"
        "  workflow_call:\n"
        "concurrency:\n"
        "  group: reusable-${{ github.ref }}\n"
        "jobs:\n"
        "  gitleaks:\n"
        "    name: gitleaks\n"
        "    timeout-minutes: 5\n"
        "    steps:\n"
        "      - uses: " + PINNED + "\n"
    )

    def policy(self, required="pr-size"):
        return {"required_jobs": [required], "events": {"pull_request": {"accepted_skips": []}},
                "state_dependent_gates": []}

    def run_files(self, files, contract, required="pr-size"):
        with tempfile.TemporaryDirectory() as d:
            d_path = Path(d)
            (d_path / "policy.json").write_text(json.dumps(self.policy(required)), encoding="utf-8")
            (d_path / "host-contract.json").write_text(json.dumps(contract), encoding="utf-8")
            args = []
            for name, text in files.items():
                (d_path / name).parent.mkdir(parents=True, exist_ok=True)
                (d_path / name).write_text(text, encoding="utf-8")
                args += ["--workflow", name]
            out = subprocess.run([sys.executable, str(GATE), *args,
                                  "--required-jobs", "policy.json", "--base-branch", "main",
                                  "--host-contract", "host-contract.json"],
                                 capture_output=True, text=True, cwd=str(d_path))
        return out.returncode, out.stdout + out.stderr

    def test_the_reusable_publishes_the_composite_name_and_enters_the_hr30_scope(self):
        """(a) El nombre compuesto se reconoce y el job que lo publica entra en el
        alcance de HR-30: el reusable dispara en `schedule` y eso es hallazgo."""
        code, out = self.run_files({
            ".github/workflows/pr-metadata.yml": self.CALLER,
            ".github/workflows/pr-size.yml": self.REUSABLE,
        }, contract_of("pr-size / pr-size"))
        self.assertEqual(code, 1, out)
        self.assertIn("publishes its check name on 'schedule'", out)
        self.assertNotIn("but no analyzed workflow job publishes it", out)

    def test_the_control_keeps_todays_finding(self):
        """(b) Control: el mismo nombre publicado por un job propio del workflow
        conserva el finding de HR-30 de hoy."""
        code, out = self.run_files({".github/workflows/plain.yml": self.PLAIN},
                                   contract_of("pr-size / pr-size"))
        self.assertEqual(code, 1, out)
        self.assertIn("publishes its check name on 'schedule'", out)
        self.assertNotIn("but no analyzed workflow job publishes it", out)

    def test_a_declared_name_nobody_publishes_is_still_a_finding(self):
        """(c) Fail closed: un nombre declarado que nada publica sigue siendo hallazgo."""
        code, out = self.run_files({
            ".github/workflows/pr-metadata.yml": self.CALLER,
            ".github/workflows/pr-size.yml": self.REUSABLE,
        }, contract_of("pr-size / pr-size", "lint / lint"))
        self.assertEqual(code, 1, out)
        self.assertIn("requires check 'lint / lint'", out)
        self.assertNotIn("requires check 'pr-size / pr-size'", out)

    def test_the_reusable_outside_the_analysed_files_publishes_nothing(self):
        """(d) Fail closed: sin el workflow reusable entre los ficheros analizados no
        se puede probar ningún nombre compuesto; el declarado sigue sin publicarse."""
        code, out = self.run_files({".github/workflows/pr-metadata.yml": self.CALLER},
                                   contract_of("pr-size / pr-size"))
        self.assertEqual(code, 1, out)
        self.assertIn("requires check 'pr-size / pr-size'", out)
        self.assertIn("pass the workflow that does", out)
        self.assertNotIn("publishes its check name on 'schedule'", out)

    def test_a_remote_reusable_publishes_nothing(self):
        """(e) Fail closed: un `uses:` remoto no está entre los ficheros analizados."""
        caller = self.CALLER.replace("uses: ./.github/workflows/pr-size.yml",
                                     "uses: otro/repo/.github/workflows/pr-size.yml@" + "a" * 40)
        code, out = self.run_files({".github/workflows/pr-metadata.yml": caller},
                                   contract_of("pr-size / pr-size"))
        self.assertEqual(code, 1, out)
        self.assertIn("requires check 'pr-size / pr-size'", out)

    def test_the_composite_name_keeps_the_wildcard_of_the_caller(self):
        """(f) La forma medida en access2web-blueprint: `security (pull_request) /
        gitleaks`; el `${{ }}` del nombre del llamador sigue siendo comodín."""
        code, out = self.run_files({
            ".github/workflows/pr-metadata.yml": self.SECURITY_CALLER,
            ".github/workflows/security.yml": self.SECURITY_REUSABLE,
        }, contract_of("security (pull_request) / gitleaks"), required="security")
        self.assertEqual(code, 0, out)
        self.assertNotIn("requires check 'security (pull_request) / gitleaks'", out)


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


class TestCancelInProgress(GateCase):
    """#978: a run cancelled after its jobs exist leaves a `cancelled` check run on the head SHA,
    and a cancelled required check blocks the merge even when a later run of the same SHA is green.
    On `pull_request` only a new commit (`synchronize`) may cancel: its victim is the old SHA."""

    def wf(self, concurrency, on="on:\n  pull_request:\n"):
        return yml(on=on).replace("concurrency:\n  group: wf-${{ github.ref }}\n", concurrency)

    def test_literal_true_on_a_pull_request_workflow_is_a_finding(self):
        text = self.wf("concurrency:\n  group: wf\n  cancel-in-progress: true\n")
        code, out = self.run_gate(text)
        self.assertEqual(code, 1, out)
        self.assertIn("cancel-in-progress", out)
        self.assertIn("synchronize", out)

    def test_string_true_is_a_finding_too(self):
        text = self.wf("concurrency:\n  group: wf\n  cancel-in-progress: 'true'\n")
        self.assertEqual(self.run_gate(text)[0], 1)

    def test_a_job_level_true_is_a_finding(self):
        text = self.wf("concurrency:\n  group: wf\n").replace(
            "    timeout-minutes: 5\n",
            "    timeout-minutes: 5\n    concurrency:\n      group: j\n      cancel-in-progress: true\n")
        code, out = self.run_gate(text)
        self.assertEqual(code, 1, out)
        self.assertIn("job 'build'", out)

    def test_false_passes(self):
        text = self.wf("concurrency:\n  group: wf\n  cancel-in-progress: false\n")
        self.assertEqual(self.run_gate(text)[0], 0)

    def test_an_expression_that_names_synchronize_passes(self):
        text = self.wf("concurrency:\n  group: wf\n"
                       "  cancel-in-progress: ${{ github.event.action == 'synchronize' }}\n")
        self.assertEqual(self.run_gate(text)[0], 0)

    def test_an_expression_that_does_not_name_synchronize_is_a_finding(self):
        text = self.wf("concurrency:\n  group: wf\n"
                       "  cancel-in-progress: ${{ github.event_name == 'pull_request' }}\n")
        self.assertEqual(self.run_gate(text)[0], 1)

    def test_without_a_pull_request_trigger_true_is_not_judged(self):
        import importlib.util
        spec = importlib.util.spec_from_file_location("cwp", Path(__file__).resolve().parents[1] / "check_workflow_policy.py")
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        doc = mod.parse_workflow(self.wf("concurrency:\n  group: wf\n  cancel-in-progress: true\n",
                                         on="on:\n  push:\n    branches: [main]\n"), "wf.yml")
        self.assertEqual(mod.cancel_in_progress_findings("wf.yml", doc), [])

    def test_absent_cancel_in_progress_passes(self):
        self.assertEqual(self.run_gate(yml())[0], 0)


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

class TestInstanceIsolatedGroups(GateCase):
    """HR-73: un gate que corre en dos instancias (su propio disparador y la llamada
    `workflow_call` de otro workflow) no puede compartir el grupo de concurrencia: GitHub
    guarda una sola petición pendiente por grupo y cancela la pendiente cuando llega otra,
    así que una instancia pierde su veredicto y el PR queda rojo sin defecto
    (BP#1371, run `ci` 40084048493 intento 1: job `pr-gate / pr-gate` cancelado a los 0 s)."""

    PAIR_POLICY = {"required_jobs": ["pr-gate"], "events": {"pull_request": {"accepted_skips": []}},
                   "state_dependent_gates": []}

    CALLER = ("name: ci\non:\n  pull_request:\nconcurrency:\n  group: wf-${{ github.ref }}\n"
              "jobs:\n  call:\n    uses: ./.github/workflows/callee.yml\n")

    def callee(self, group, direct=True, inputs=""):
        on = "on:\n  pull_request:\n  workflow_call:\n" if direct else "on:\n  workflow_call:\n"
        return (on + inputs + "jobs:\n  pr-gate:\n    timeout-minutes: 5\n"
                "    concurrency:\n      group: " + group + "\n      cancel-in-progress: false\n"
                "    steps:\n      - " + PINNED + "\n")

    def pair(self, caller, callee, policy=None):
        with tempfile.TemporaryDirectory() as d:
            c, w, p = Path(d, "caller.yml"), Path(d, "callee.yml"), Path(d, "policy.json")
            c.write_text(caller, encoding="utf-8")
            w.write_text(callee, encoding="utf-8")
            p.write_text(json.dumps(policy or self.PAIR_POLICY), encoding="utf-8")
            out = subprocess.run([sys.executable, str(GATE), "--workflow", str(c),
                                  "--workflow", str(w), "--required-jobs", str(p),
                                  "--base-branch", "main"], capture_output=True, text=True)
        return out.returncode, out.stdout + out.stderr

    def test_group_shared_by_the_two_instances_is_a_finding(self):
        code, out = self.pair(
            self.CALLER,
            self.callee("gate-pr-gate-${{ github.event.pull_request.number || github.ref }}"))
        self.assertEqual(code, 1, out)
        self.assertIn("pr-gate", out)
        self.assertIn("gate-pr-gate-", out)
        self.assertIn("HR-73", out)

    def test_github_workflow_in_the_group_separates_the_instances(self):
        code, out = self.pair(
            self.CALLER,
            self.callee("gate-pr-gate-${{ github.workflow }}-${{ github.event.pull_request.number }}"))
        self.assertEqual(code, 0, out)

    def test_github_run_id_in_the_group_separates_the_instances(self):
        code, out = self.pair(self.CALLER, self.callee("gate-pr-gate-${{ github.run_id }}"))
        self.assertEqual(code, 0, out)

    def test_a_single_instance_is_clean(self):
        # Sin disparador propio el reusable solo corre como llamada: una sola instancia.
        # El required job pasa a ser el del caller (el reusable no publica nada de PR).
        policy = {"required_jobs": ["call"], "events": {"pull_request": {"accepted_skips": []}},
                  "state_dependent_gates": []}
        code, out = self.pair(
            self.CALLER,
            self.callee("gate-pr-gate-${{ github.event.pull_request.number || github.ref }}",
                        direct=False),
            policy=policy)
        self.assertEqual(code, 0, out)

    def test_a_reusable_workflow_outside_the_analysed_set_is_clean(self):
        # La segunda instancia no se puede probar si el reusable no está entre los
        # workflows analizados: el gate no inventa.
        policy = {"required_jobs": ["call"], "events": {"pull_request": {"accepted_skips": []}},
                  "state_dependent_gates": []}
        code, out = self.run_gate(self.CALLER, policy=policy)
        self.assertEqual(code, 0, out)

    def test_two_calls_from_the_same_caller_are_a_finding(self):
        # `github.workflow` vale el caller: dos llamadas del mismo caller comparten valor.
        caller = ("name: ci\non:\n  pull_request:\nconcurrency:\n  group: wf-${{ github.ref }}\n"
                  "jobs:\n  call:\n    uses: ./.github/workflows/callee.yml\n"
                  "  other:\n    uses: ./.github/workflows/callee.yml\n")
        code, out = self.pair(
            caller,
            self.callee("gate-pr-gate-${{ github.workflow }}-${{ github.event.pull_request.number }}"))
        self.assertEqual(code, 1, out)
        self.assertIn("HR-73", out)

    def test_an_input_the_caller_fills_separates_the_instances(self):
        caller = ("name: ci\non:\n  pull_request:\nconcurrency:\n  group: wf-${{ github.ref }}\n"
                  "jobs:\n  call:\n    uses: ./.github/workflows/callee.yml\n"
                  "    with:\n      instance: ci\n")
        inputs = "    inputs:\n      instance:\n        required: false\n"
        code, out = self.pair(
            caller,
            self.callee("gate-pr-gate-${{ inputs.instance || github.event.pull_request.number }}",
                        inputs=inputs))
        self.assertEqual(code, 0, out)

    def test_an_input_the_caller_leaves_empty_is_a_finding(self):
        inputs = "    inputs:\n      instance:\n        required: false\n"
        code, out = self.pair(
            self.CALLER,
            self.callee("gate-pr-gate-${{ inputs.instance || github.event.pull_request.number }}",
                        inputs=inputs))
        self.assertEqual(code, 1, out)
        self.assertIn("HR-73", out)


class TestHR30ApiPublishedCheckNames(RepoCase):
    """A check the job POSTs through the check-runs API counts as published
    under its literal CONTEXT (DysTelefonica/team-skills#1313, G3). The shape is the
    canonical pr-contract template: a pull_request_target job, a name the host contract
    requires, and no job display name that equals it."""

    CONTRACT = {"contract_version": 1,
                "required_checks": [{"name": "pr-contract / pr-contract", "class": "runner-enforced"}]}
    POLICY_PRT = {"required_jobs": ["publish"],
                  "events": {"pull_request_target": {"accepted_skips": []}},
                  "state_dependent_gates": []}

    @staticmethod
    def workflow(context="\"pr-contract / pr-contract\"", on="on:\n  pull_request_target:\n",
                 publish=True):
        step = ""
        if publish:
            step = ("      - name: Publish the verdict as a check run on the PR head\n"
                    "        env:\n"
                    f"          CONTEXT: {context}\n"
                    "        run: |\n"
                    "          gh api --method POST \"repos/$GITHUB_REPOSITORY/check-runs\" --input x.json\n")
        else:
            step = "      - name: Verdict\n        run: |\n          echo ok\n"
        return (on + "concurrency:\n  group: wf-${{ github.ref }}\njobs:\n  publish:\n"
                "    name: publish\n    timeout-minutes: 5\n    steps:\n"
                "      - " + PINNED + "\n" + step)

    def run_prt(self, text):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            w, p, c = root / "wf.yml", root / "policy.json", root / "host-contract.json"
            w.write_text(text, encoding="utf-8")
            p.write_text(json.dumps(self.POLICY_PRT), encoding="utf-8")
            c.write_text(json.dumps(self.CONTRACT), encoding="utf-8")
            out = subprocess.run([sys.executable, str(GATE), "--workflow", str(w),
                                  "--required-jobs", str(p), "--base-branch", "main",
                                  "--repo-root", str(root), "--host-contract", str(c)],
                                 capture_output=True, text=True)
        return out.returncode, out.stdout + out.stderr

    def test_literal_context_publishes_the_declared_name(self):
        code, out = self.run_prt(self.workflow())
        self.assertEqual(code, 0, out)
        self.assertNotIn("no analyzed workflow job publishes it", out)

    def test_without_the_check_runs_step_the_name_is_still_a_finding(self):
        code, out = self.run_prt(self.workflow(publish=False))
        self.assertEqual(code, 1, out)
        self.assertIn("requires check 'pr-contract / pr-contract'", out)

    def test_non_literal_context_declares_no_name(self):
        code, out = self.run_prt(self.workflow(context="\"${{ matrix.check }}\""))
        self.assertEqual(code, 1, out)
        self.assertIn("requires check 'pr-contract / pr-contract'", out)

    def test_api_published_name_on_schedule_is_in_hr30_scope(self):
        code, out = self.run_prt(self.workflow(on="on:\n  pull_request_target:\n  schedule:\n    - cron: '0 3 * * *'\n"))
        self.assertEqual(code, 1, out)
        self.assertIn("does not evaluate the PR", out)


if __name__ == "__main__":
    unittest.main(verbosity=2)


GATE_OUTPUTS = ("    outputs:\n"
                "      lockfile_scope_changed: ${{ steps.lanes.outputs.lockfile_scope_changed }}\n")


def marker_pr_wf(if_expr, gate_outputs=GATE_OUTPUTS, gate_uses=None):
    """PR sin paths: un gate de primer nivel que publica el marcador y un pr_job
    `security` cuyo `if:` lo referencia (#1192)."""
    gate = "  gate:\n    name: gate\n    timeout-minutes: 5\n    concurrency: g\n"
    if gate_uses:
        gate += f"    uses: {gate_uses}\n"
    else:
        gate += gate_outputs + "    steps:\n      - uses: " + SCHED_PINNED + "\n        id: lanes\n"
    sec = ("  security:\n    name: security\n    timeout-minutes: 5\n    concurrency: j\n"
           + (f"    if: {if_expr}\n" if if_expr is not None else "")
           + "    steps:\n      - uses: " + SCHED_PINNED + "\n      - run: pip-audit\n")
    return "on:\n  pull_request:\njobs:\n" + GOOD_BUILD + gate + sec


PR_GATE_CALLED = """on:
  workflow_call:
    outputs:
      lockfile_scope_changed:
        value: ${{ jobs.pr-gate.outputs.lockfile_scope_changed }}
jobs:
  pr-gate:
    name: pr-gate
    timeout-minutes: 5
    concurrency: g
    steps:
      - uses: """ + SCHED_PINNED + """
"""

MARKER_IF = "needs.gate.outputs.lockfile_scope_changed != 'false'"


class TestHR39MarkerIfEquivalentToPaths(unittest.TestCase):
    """#1192: un pr_job cuyo `if:` referencia el marcador de un job de primer
    nivel que lo publica cuenta como filtro de ruta. Sin marcador, sigue el hallazgo."""

    def run_files(self, files):
        pol = dict(REQUIRED_POLICY)
        pol["state_dependent_gates"] = [GOOD_STATE_ENTRY]
        with tempfile.TemporaryDirectory() as d:
            d_path = Path(d)
            (d_path / "policy.json").write_text(json.dumps(pol), encoding="utf-8")
            args = []
            for name, text in files.items():
                (d_path / name).parent.mkdir(parents=True, exist_ok=True)
                (d_path / name).write_text(text, encoding="utf-8")
                args += ["--workflow", name]
            out = subprocess.run([sys.executable, str(GATE), *args,
                                  "--required-jobs", "policy.json", "--base-branch", "main"],
                                 capture_output=True, text=True, cwd=str(d_path))
        return out.returncode, out.stdout + out.stderr

    def sched(self):
        return sched_wf()

    def test_marker_if_on_first_level_gate_replaces_paths(self):
        """(f) if: con marcador de un gate que lo publica: sin hallazgo de paths."""
        code, out = self.run_files({"wf_pr.yml": marker_pr_wf(MARKER_IF),
                                    "wf_sched.yml": self.sched()})
        self.assertEqual(code, 0, out)
        self.assertNotIn("HR-39", out)

    def test_marker_if_on_reusable_gate_replaces_paths(self):
        """(g) el gate es un workflow reutilizable local que declara el marcador
        en workflow_call.outputs: vale igual."""
        pr = marker_pr_wf(MARKER_IF, gate_outputs=None,
                          gate_uses="./.github/workflows/pr-gate.yml")
        code, out = self.run_files({"wf_pr.yml": pr, "wf_sched.yml": self.sched(),
                                    ".github/workflows/pr-gate.yml": PR_GATE_CALLED})
        self.assertEqual(code, 0, out)
        self.assertNotIn("HR-39", out)

    def test_marker_if_on_unknown_gate_is_finding(self):
        """(h) el marcador apunta a un job que no existe: sigue sin paths."""
        code, out = self.run_files({
            "wf_pr.yml": marker_pr_wf("needs.nope.outputs.lockfile_scope_changed != 'false'"),
            "wf_sched.yml": self.sched()})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-39", out)
        self.assertIn("paths", out)

    def test_pr_job_without_if_is_finding(self):
        """(i) pr_job sin if: ni paths: sigue siendo hallazgo."""
        code, out = self.run_files({"wf_pr.yml": marker_pr_wf(None),
                                    "wf_sched.yml": self.sched()})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-39", out)
        self.assertIn("paths", out)

    def test_if_without_marker_comparison_is_finding(self):
        """(j) un if: que no compara un marcador de salida (solo el evento) no cuenta."""
        code, out = self.run_files({
            "wf_pr.yml": marker_pr_wf("github.event_name == 'pull_request'"),
            "wf_sched.yml": self.sched()})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-39", out)

    def test_marker_not_published_by_gate_is_finding(self):
        """(k) el gate existe pero no publica el marcador que el if: referencia."""
        code, out = self.run_files({
            "wf_pr.yml": marker_pr_wf(MARKER_IF, gate_outputs="    outputs:\n      otro: x\n"),
            "wf_sched.yml": self.sched()})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-39", out)


CALLED_AUDIT = """on:
  workflow_call:
jobs:
  audit:
    name: audit
    timeout-minutes: 5
    concurrency: a
    steps:
      - uses: """ + SCHED_PINNED + """
      - run: pip-audit --strict
"""

CALLED_NO_COMMAND = CALLED_AUDIT.replace("pip-audit --strict", "echo nada")


def pr_uses_wf(uses):
    """PR con paths y un pr_job `security` que llama a un workflow reutilizable (#1196)."""
    return ("on:\n  pull_request:\n    paths: [uv.lock, requirements*.txt]\njobs:\n"
            + GOOD_BUILD
            + f"  security:\n    uses: {uses}\n    concurrency: s\n")


class TestHR39ReusableCommand(unittest.TestCase):
    """#1196: la comprobacion de comando sigue un `uses:` local hacia el workflow
    reutilizable llamado; un `uses:` externo sigue sin resolverse."""

    def run_files(self, files):
        pol = dict(REQUIRED_POLICY)
        pol["state_dependent_gates"] = [GOOD_STATE_ENTRY]
        with tempfile.TemporaryDirectory() as d:
            d_path = Path(d)
            (d_path / "policy.json").write_text(json.dumps(pol), encoding="utf-8")
            args = []
            for name, text in files.items():
                (d_path / name).parent.mkdir(parents=True, exist_ok=True)
                (d_path / name).write_text(text, encoding="utf-8")
                args += ["--workflow", name]
            out = subprocess.run([sys.executable, str(GATE), *args,
                                  "--required-jobs", "policy.json", "--base-branch", "main"],
                                 capture_output=True, text=True, cwd=str(d_path))
        return out.returncode, out.stdout + out.stderr

    def test_reusable_call_that_runs_the_command_is_clean(self):
        """(a) el workflow llamado ejecuta el comando: sin hallazgo HR-39."""
        code, out = self.run_files({
            "wf_pr.yml": pr_uses_wf("./.github/workflows/audit.yml"),
            ".github/workflows/audit.yml": CALLED_AUDIT,
            "wf_sched.yml": sched_wf()})
        self.assertNotIn("HR-39", out)

    def test_reusable_call_without_the_command_names_the_called_workflow(self):
        """(b) el workflow llamado no ejecuta el comando: el hallazgo nombra ese workflow."""
        code, out = self.run_files({
            "wf_pr.yml": pr_uses_wf("./.github/workflows/audit.yml"),
            ".github/workflows/audit.yml": CALLED_NO_COMMAND,
            "wf_sched.yml": sched_wf()})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-39", out)
        self.assertIn(".github/workflows/audit.yml", out)

    def test_reusable_call_to_an_unlisted_workflow_names_the_path(self):
        """(c) el workflow llamado no está entre los --workflow de entrada: hallazgo que lo nombra."""
        code, out = self.run_files({
            "wf_pr.yml": pr_uses_wf("./.github/workflows/audit.yml"),
            "wf_sched.yml": sched_wf()})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-39", out)
        self.assertIn(".github/workflows/audit.yml", out)

    def test_external_reusable_call_stays_a_finding(self):
        """(d) un `uses:` externo no se resuelve: hallazgo (sin cambio de comportamiento)."""
        code, out = self.run_files({
            "wf_pr.yml": pr_uses_wf("otro/repo/.github/workflows/audit.yml@" + "a" * 40),
            "wf_sched.yml": sched_wf()})
        self.assertEqual(code, 1, out)
        self.assertIn("HR-39", out)
        self.assertIn("ejecuta", out)


class TestHR79NeedsOutputsFromDirectDependencies(GateCase):
    """HR-79: `needs.<job>.outputs` solo existe si el job es dependencia DIRECTA.

    El contexto `needs` de un job expone los outputs de sus dependencias
    directas: leer el de una transitiva devuelve vacío, sin error, y el
    workflow queda en verde con el valor equivocado (#1277). El caso real es
    el de access2web-blueprint: `publish` leía el canal de `preflight`, que
    solo era dependencia transitiva, y una RC salía como release estable.
    """

    def workflow(self, read, where="env", needs="[e2e]"):
        """Workflow de cuatro jobs: `build` (requerido), `preflight`, `e2e` y `publish`."""
        head = f"    if: needs.{read}.outputs.kind == 'rc'\n" if where == "if" else ""
        step = {
            "env": f"      - run: echo ok\n        env:\n          KIND: ${{{{ needs.{read}.outputs.kind }}}}\n",
            "with": f"      - uses: {PINNED}\n        with:\n          kind: ${{{{ needs.{read}.outputs.kind }}}}\n",
            "if": "      - run: echo ok\n",
        }[where]
        return (
            "on:\n  pull_request:\n"
            "concurrency:\n  group: wf-${{ github.ref }}\n"
            "jobs:\n"
            "  build:\n"
            "    name: build (${{ github.event_name }})\n"
            "    timeout-minutes: 5\n"
            "    steps:\n"
            f"      - {PINNED}\n"
            "  preflight:\n"
            "    name: preflight\n"
            "    timeout-minutes: 5\n"
            "    outputs:\n"
            "      kind: ${{ steps.id.outputs.kind }}\n"
            "    steps:\n"
            f"      - {PINNED}\n"
            "  e2e:\n"
            "    name: e2e\n"
            "    needs: preflight\n"
            "    timeout-minutes: 5\n"
            "    steps:\n"
            f"      - {PINNED}\n"
            "  publish:\n"
            "    name: publish\n"
            f"{head}"
            f"    needs: {needs}\n"
            "    timeout-minutes: 5\n"
            "    steps:\n"
            f"{step}"
        )

    def test_output_of_a_direct_dependency_is_clean(self):
        """`e2e` es dependencia directa de `publish`: su output llega."""
        code, out = self.run_gate(self.workflow("e2e"))
        self.assertEqual(code, 0, out)
        self.assertNotIn("HR-79", out)

    def test_output_of_a_transitive_dependency_is_a_finding(self):
        """`preflight` solo es transitiva para `publish`: el valor llega vacío."""
        code, out = self.run_gate(self.workflow("preflight"))
        self.assertEqual(code, 1, out)
        self.assertIn("HR-79", out)
        self.assertIn("preflight", out)
        self.assertIn("publish", out)

    def test_the_rule_does_not_depend_on_where_the_job_reads_it(self):
        """`if:`, `env:` y `with:`: la regla no depende de dónde se lea."""
        for where in ("if", "env", "with"):
            with self.subTest(where=where):
                code, out = self.run_gate(self.workflow("preflight", where=where))
                self.assertEqual(code, 1, out)
                self.assertIn("HR-79", out)


class PullRequestCatalogSecretTests(GateCase):
    """#1270: el token del catálogo solo en jobs que no pueden correr en un PR.

    En `pull_request` el workflow, las acciones locales y el sello vienen de la
    cabeza del PR: un job que puede correr ahí puede leer el token. Solo
    `pull_request_target` (código de la base), `schedule` y `workflow_dispatch`
    pueden tenerlo.
    """

    SECRET = "secrets.TEAM_SKILLS_READ_TOKEN"

    def wf(self, on="on:\n  pull_request:\n", job="build", cond=None, secret=True):
        s = (f"{on}concurrency:\n  group: wf-${{{{ github.ref }}}}\njobs:\n  {job}:\n"
             f"    name: {job} (${{{{ github.event_name }}}})\n    timeout-minutes: 5\n")
        if cond:
            s += f"    if: {cond}\n"
        s += "    steps:\n"
        if secret:
            s += ("      - name: Resolve the pattern\n"
                  "        uses: ./.github/actions/pattern-at-pin\n"
                  "        with:\n"
                  f"          token: ${{{{ {self.SECRET} }}}}\n")
        else:
            s += "      - run: echo ok\n"
        return s

    def run_workflows(self, files: dict, required=("build",)):
        """Analiza varios workflows a la vez (para el grafo de los reutilizables)."""
        with tempfile.TemporaryDirectory() as d:
            args = []
            for name, text in files.items():
                p = Path(d, name)
                p.parent.mkdir(parents=True, exist_ok=True)
                p.write_text(text, encoding="utf-8")
                args += ["--workflow", str(p)]
            policy = Path(d, "policy.json")
            policy.write_text(json.dumps({"required_jobs": list(required),
                                          "events": {"pull_request": {"accepted_skips": []}},
                                          "state_dependent_gates": []}), encoding="utf-8")
            out = subprocess.run([sys.executable, str(GATE), *args,
                                  "--required-jobs", str(policy), "--base-branch", "main"],
                                 capture_output=True, text=True)
        return out.returncode, out.stdout + out.stderr

    def test_a_pull_request_job_that_reads_the_secret_is_a_finding(self):
        code, out = self.run_gate(self.wf())
        self.assertEqual(code, 1, out)
        self.assertIn(self.SECRET, out)
        self.assertIn("#1270", out)

    def test_the_same_job_in_pull_request_target_is_clean(self):
        code, out = self.run_gate(self.wf(on="on:\n  pull_request_target:\n"))
        self.assertEqual(code, 0, out)

    def test_schedule_and_dispatch_are_clean(self):
        """Un workflow programado puede tener el token: no evalúa el PR."""
        for on in ("on:\n  schedule:\n    - cron: \"0 4 * * 1\"\n", "on:\n  workflow_dispatch:\n"):
            with self.subTest(on=on):
                code, out = self.run_workflows({
                    "wf.yml": self.wf(on=on, job="other"),
                    "pr.yml": self.wf(job="build", secret=False),
                }, required=("build",))
                self.assertEqual(code, 0, out)

    def test_a_job_that_cannot_run_on_a_pr_is_clean(self):
        """La condición del job lo excluye del PR: no lo puede leer."""
        code, out = self.run_gate(self.wf(cond="github.event_name != 'pull_request'"))
        self.assertEqual(code, 0, out)

    def test_a_workflow_without_the_secret_is_clean(self):
        code, out = self.run_gate(self.wf(secret=False))
        self.assertEqual(code, 0, out)

    def test_a_reusable_workflow_called_from_a_pr_inherits_the_event(self):
        """El workflow reusable corre con los eventos de quien lo llama."""
        caller = ("on:\n  pull_request:\nconcurrency:\n  group: wf-${{ github.ref }}\njobs:\n"
                  "  call:\n    name: call\n    timeout-minutes: 5\n"
                  "    uses: ./.github/workflows/callee.yml\n")
        code, out = self.run_workflows({
            ".github/workflows/callee.yml": self.wf(on="on:\n  workflow_call:\n", job="gate"),
            ".github/workflows/caller.yml": caller,
        }, required=("call",))
        self.assertEqual(code, 1, out)
        self.assertIn(self.SECRET, out)

    def test_a_reusable_workflow_called_only_from_a_target_is_clean(self):
        caller = ("on:\n  pull_request_target:\nconcurrency:\n  group: wf-${{ github.ref }}\njobs:\n"
                  "  call:\n    name: call\n    timeout-minutes: 5\n"
                  "    uses: ./.github/workflows/callee.yml\n")
        code, out = self.run_workflows({
            ".github/workflows/callee.yml": self.wf(on="on:\n  workflow_call:\n", job="gate"),
            ".github/workflows/caller.yml": caller,
        }, required=("call",))
        self.assertEqual(code, 0, out)
