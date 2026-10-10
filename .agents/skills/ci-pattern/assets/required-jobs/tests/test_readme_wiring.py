#!/usr/bin/env python3
# ci-pattern asset tests — the wiring the README documents really runs (DysTelefonica/team-skills#161)
"""Executes the aggregator and parity steps exactly as the README tells a
consumer to write them: each ``run:`` script is lifted from the README and
run with bash, with the ``env:`` variables the same step declares.

The README is the installation contract; a wiring that only reads well is
not evidence. On ``pull_request`` the step fetches the base branch, so each
run gets a throwaway ``origin`` repository to fetch from. Stdlib only, plus
the ``bash`` and ``git`` every CI runner ships.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import textwrap
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
README = ASSET / "README.md"
EXAMPLE_POLICY = ASSET / "required-jobs.policy.example.json"
WORKFLOW_FIXTURE = Path(__file__).resolve().parent / "fixtures" / "ci-workflow.fixture.yml"

# Where the README tells the consumer to install the asset and its policy.
INSTALL_DIR = "scripts/required-jobs"
POLICY_NAME = "required-jobs.policy.json"

POLICY_JOBS = ["lint", "compile", "smoke", "unit", "full"]
BASE_BRANCH = "main"

ABSENT = object()  # sentinel: the base branch holds no policy file yet


def git(cwd: Path, *args: str) -> None:
    """Run git with a fixed identity, independent of the machine's config."""
    subprocess.run(
        ["git", "-c", "user.name=test", "-c", "user.email=test@example.invalid",
         "-c", "commit.gpgsign=false", "-c", "init.defaultBranch=" + BASE_BRANCH, *args],
        cwd=cwd, check=True, capture_output=True, text=True,
    )


def documented_step(readme_text: str, script_name: str) -> tuple[list[str], str]:
    """Return the env names and the ``run:`` script of a documented step.

    Reads the first YAML fence that invokes ``script_name``. Raises when the
    README no longer carries such a step, so the suite cannot pass on nothing.
    """
    for fence in readme_text.split("```yaml")[1:]:
        block = fence.split("```", 1)[0]
        if script_name not in block:
            continue
        lines = block.splitlines()
        env_names: list[str] = []
        script: list[str] = []
        section = None
        section_indent = 0
        for line in lines:
            stripped = line.strip()
            indent = len(line) - len(line.lstrip())
            if section and stripped and indent <= section_indent:
                section = None
            # A key may open the step ("- run: |") or follow another key.
            key = stripped[2:] if stripped.startswith("- ") else stripped
            if section == "env" and stripped:
                env_names.append(stripped.split(":", 1)[0])
            elif section == "run":
                script.append(line)
            elif key == "env:":
                section, section_indent = "env", indent
            elif key == "run: |":
                section, section_indent = "run", indent
        if not script:
            raise AssertionError(f"the README step invoking {script_name} has no 'run: |' script")
        return env_names, textwrap.dedent("\n".join(script)) + "\n"
    raise AssertionError(f"the README has no YAML step invoking {script_name}")


def needs_object(overrides: dict | None = None, outputs: dict | None = None) -> str:
    """Serialize a needs object the way ``toJSON(needs)`` does."""
    results = {job: "success" for job in POLICY_JOBS}
    results.update(overrides or {})
    needs = {job: {"result": result, "outputs": outputs or {}} for job, result in results.items()}
    return json.dumps(needs, indent=2)


def example_policy() -> dict:
    return json.loads(EXAMPLE_POLICY.read_text(encoding="utf-8"))


class DocumentedStepCase(unittest.TestCase):
    """Shared harness: lifts the step from the README and runs it."""

    SCRIPT_NAME = "check_required_jobs.py"

    @classmethod
    def setUpClass(cls) -> None:
        missing = [tool for tool in ("bash", "git") if not shutil.which(tool)]
        if missing:
            message = f"cannot execute the documented step without {missing}"
            # In CI a suite that did not run must not read as green.
            if os.environ.get("CI"):
                raise AssertionError(message)
            raise unittest.SkipTest(message)
        cls.env_names, cls.script = documented_step(
            README.read_text(encoding="utf-8"), cls.SCRIPT_NAME)

    def run_step(self, needs_json: str, event: str = "push", candidate: dict | None = None,
                 base: object = None, base_ref: str = BASE_BRANCH,
                 base_tip: dict | None = None,
                 ) -> tuple[subprocess.CompletedProcess, Path]:
        """Run the documented script inside a consumer-shaped checkout.

        ``candidate`` is the policy in the checkout (default: the example).
        ``base`` is the policy committed on the base branch of ``origin``:
        the example by default, a dict, or ``ABSENT`` for a base branch that
        predates the gate.
        """
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp, ignore_errors=True)
        origin, work, runner_temp = tmp / "origin", tmp / "work", tmp / "runner-temp"
        for directory in (origin / INSTALL_DIR, work / INSTALL_DIR, runner_temp):
            directory.mkdir(parents=True)

        (origin / "README.md").write_text("base branch\n", encoding="utf-8")
        if base is not ABSENT:
            base_policy = example_policy() if base is None else base
            (origin / INSTALL_DIR / POLICY_NAME).write_text(json.dumps(base_policy), encoding="utf-8")
        git(origin, "init")
        git(origin, "add", "-A")
        git(origin, "commit", "-m", "base")
        # El PR apunta a este commit: BASE_SHA. La punta de la rama base
        # puede avanzar después (base_tip) sin cambiar BASE_SHA.
        base_sha = subprocess.run(
            ["git", "-C", str(origin), "rev-parse", "HEAD"],
            check=True, capture_output=True, text=True).stdout.strip()
        if base_tip is not None:
            (origin / INSTALL_DIR / POLICY_NAME).write_text(json.dumps(base_tip), encoding="utf-8")
            git(origin, "add", "-A")
            git(origin, "commit", "-m", "tip moved")

        shutil.copy(ASSET / "check_required_jobs.py", work / INSTALL_DIR)
        (work / INSTALL_DIR / POLICY_NAME).write_text(
            json.dumps(example_policy() if candidate is None else candidate), encoding="utf-8")
        git(work, "init")
        git(work, "remote", "add", "origin", origin.as_uri())

        values = {"NEEDS_JSON": needs_json, "EVENT_NAME": event,
                  "BASE_REF": base_ref if event == "pull_request" else "",
                  "BASE_SHA": base_sha if event == "pull_request" else ""}
        # Every variable the step declares must be one this suite knows how
        # to provide; an unknown one means the README outgrew the test.
        self.assertEqual(sorted(self.env_names), sorted(values))
        proc = subprocess.run(
            # The default shell of a GitHub-hosted run step: bash -e.
            ["bash", "-e", "-c", self.script],
            cwd=work, env={**os.environ, **values, "RUNNER_TEMP": str(runner_temp)},
            capture_output=True, text=True,
        )
        return proc, work


class ReadmeWiringTests(DocumentedStepCase):
    """The documented step must reach the gate's verdict, pass and fail."""

    def test_script_carries_no_workflow_expression(self) -> None:
        # An expression inside run: is substituted before the shell parses it.
        self.assertNotIn("${{", self.script)

    def test_all_success_passes(self) -> None:
        proc, _ = self.run_step(needs_object())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("VERDICT: PASS (event: push)", proc.stdout)

    def test_failed_job_fails_the_step(self) -> None:
        proc, _ = self.run_step(needs_object({"unit": "failure"}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("unit (failure)", proc.stdout)

    def test_declared_skip_depends_on_the_event(self) -> None:
        needs = needs_object({"full": "skipped"})
        accepted, _ = self.run_step(needs, event="pull_request")
        self.assertEqual(accepted.returncode, 0, accepted.stdout + accepted.stderr)
        refused, _ = self.run_step(needs, event="push")
        self.assertEqual(refused.returncode, 1, refused.stdout + refused.stderr)

    def test_quotes_in_job_outputs_neither_break_nor_inject(self) -> None:
        hostile = {"note": "it's '; touch pwned; echo ' and \"$(touch pwned)\" `touch pwned`"}
        proc, tmp = self.run_step(needs_object(outputs=hostile))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertFalse((tmp / "pwned").exists(), "a job output was executed by the step")

    def test_hostile_event_name_is_data_not_code(self) -> None:
        proc, tmp = self.run_step(needs_object(), event="push'; touch pwned; '")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("is not declared in the policy", proc.stdout)
        self.assertFalse((tmp / "pwned").exists(), "the event name was executed by the step")


class ReadmeBaseComparisonTests(DocumentedStepCase):
    """On pull_request the documented step compares against the base branch."""

    def relaxed(self) -> dict:
        policy = example_policy()
        policy["events"]["push"]["accepted_skips"] = ["full"]
        return policy

    def test_pull_request_performs_the_base_comparison(self) -> None:
        proc, _ = self.run_step(needs_object(), event="pull_request")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("base comparison: performed, no relaxation detected", proc.stdout)

    def test_pull_request_relaxing_the_policy_fails_the_step(self) -> None:
        proc, _ = self.run_step(needs_object(), event="pull_request", candidate=self.relaxed())
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("event 'push': accepted skip 'full' added", proc.stdout)
        self.assertIn("REFUSED", proc.stdout)

    def test_pull_request_relaxing_with_its_field_passes(self) -> None:
        candidate = self.relaxed()
        candidate["relaxation"] = {"reason": "full runs nightly", "reference": "example/repo#12",
                                   "issue": 12}
        proc, _ = self.run_step(needs_object(), event="pull_request", candidate=candidate)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("covered by relaxation", proc.stdout)

    def test_invalid_policy_on_the_base_branch_fails_the_step(self) -> None:
        proc, _ = self.run_step(needs_object(), event="pull_request",
                                base={"required_jobs": POLICY_JOBS})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("base policy file", proc.stderr)

    def test_base_branch_without_policy_declares_the_comparison_undone(self) -> None:
        # First adoption: the PR that installs the gate has no base copy.
        proc, _ = self.run_step(needs_object(), event="pull_request", base=ABSENT)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("base comparison: NOT performed", proc.stdout)

    def test_base_branch_that_cannot_be_fetched_fails_the_step(self) -> None:
        proc, _ = self.run_step(needs_object(), event="pull_request", base_ref="no-such-branch")
        self.assertNotEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("VERDICT", proc.stdout)

    def test_comparison_uses_the_pinned_base_sha_not_the_tip(self) -> None:
        # La punta de la rama base ya relajó la política; el SHA base del PR
        # no. La relajación del candidato se mide contra el SHA base: si el
        # paso comparara contra la punta, pasaría de silencio.
        tip_policy = example_policy()
        tip_policy["events"]["push"]["accepted_skips"] = ["full"]
        proc, _ = self.run_step(needs_object(), event="pull_request",
                                candidate=self.relaxed(),
                                base_tip=tip_policy)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("REFUSED", proc.stdout)

    def test_other_events_declare_the_comparison_undone(self) -> None:
        proc, _ = self.run_step(needs_object(), event="push", candidate=self.relaxed())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("base comparison: NOT performed", proc.stdout)


class ReadmeParityStepTests(DocumentedStepCase):
    """The documented parity step runs against the consumer's own workflow."""

    SCRIPT_NAME = "check_parity.py"
    WORKFLOW = ".github/workflows/ci.yml"  # the path the README step names

    def run_parity_step(self, workflow_text: str) -> subprocess.CompletedProcess:
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp, ignore_errors=True)
        install = tmp / INSTALL_DIR
        install.mkdir(parents=True)
        for name in ("check_parity.py", "check_required_jobs.py"):
            shutil.copy(ASSET / name, install)
        shutil.copy(EXAMPLE_POLICY, install / POLICY_NAME)
        workflow = tmp / self.WORKFLOW
        workflow.parent.mkdir(parents=True)
        workflow.write_text(workflow_text, encoding="utf-8")
        self.assertEqual(self.env_names, [])
        return subprocess.run(["bash", "-e", "-c", self.script], cwd=tmp,
                              capture_output=True, text=True)

    def test_workflow_in_parity_passes_the_step(self) -> None:
        proc = self.run_parity_step(WORKFLOW_FIXTURE.read_text(encoding="utf-8"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("PARITY: OK", proc.stdout)

    def test_job_added_without_policy_fails_the_step(self) -> None:
        grown = WORKFLOW_FIXTURE.read_text(encoding="utf-8").replace(
            "  required:\n", "  e2e:\n    runs-on: ubuntu-latest\n  required:\n")
        proc = self.run_parity_step(grown)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("workflow job 'e2e' is not in the policy", proc.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
