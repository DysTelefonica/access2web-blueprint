#!/usr/bin/env python3
# ci-pattern asset tests — parity between a workflow and the policy (DysTelefonica/team-skills#132, #161)
"""Behavioural suite for ``check_parity.py`` (HR-29).

The check is run as a subprocess, the way a consumer runs it against its own
workflow: on the committed fixtures, which must be in parity with the example
policy, and on workflows that break parity or that the parser cannot read.
Every mismatch and every doubt asserts a NON-ZERO exit. Stdlib only.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "check_parity.py"
EXAMPLE_POLICY = ASSET / "required-jobs.policy.example.json"
FIXTURES = Path(__file__).resolve().parent / "fixtures"
FIXTURE = FIXTURES / "ci-workflow.fixture.yml"
BLOCK_FIXTURE = FIXTURES / "ci-workflow-block-needs.fixture.yml"

AGGREGATOR = "required"
POLICY_JOBS = ["lint", "compile", "smoke", "unit", "full"]


def run_parity(workflow: Path | str, policy: Path | dict | None = None,
               aggregator: str = AGGREGATOR) -> subprocess.CompletedProcess:
    """Run check_parity.py. ``workflow`` is a path or the YAML text itself;
    ``policy`` is a path, a policy dict, or None for the example policy."""
    with tempfile.TemporaryDirectory() as tmp:
        if isinstance(workflow, str):
            workflow_path = Path(tmp) / "ci.yml"
            workflow_path.write_text(textwrap.dedent(workflow), encoding="utf-8")
        else:
            workflow_path = workflow
        if policy is None:
            policy_path = EXAMPLE_POLICY
        elif isinstance(policy, dict):
            policy_path = Path(tmp) / "policy.json"
            policy_path.write_text(json.dumps(policy), encoding="utf-8")
        else:
            policy_path = policy
        return subprocess.run(
            [sys.executable, str(SCRIPT), "--workflow", str(workflow_path),
             "--aggregator", aggregator, "--policy", str(policy_path)],
            capture_output=True, text=True,
        )


def policy_for(jobs: list[str]) -> dict:
    return {"required_jobs": jobs, "events": {"push": {"accepted_skips": []}}}


def workflow_with(needs_yaml: str, extra_jobs: str = "") -> str:
    """A three-job workflow whose aggregator carries ``needs_yaml`` verbatim."""
    body = textwrap.dedent("""\
        name: ci
        on: [push]
        jobs:
          lint:
            runs-on: ubuntu-latest
          unit:
            runs-on: ubuntu-latest
        """)
    body += textwrap.indent(textwrap.dedent(extra_jobs), "  ")
    body += "  required:\n    if: always()\n"
    body += textwrap.indent(textwrap.dedent(needs_yaml), "    ")
    return body + "    runs-on: ubuntu-latest\n"


class ParityTests(unittest.TestCase):
    """Workflow jobs, aggregator needs and policy must be one set."""

    def test_fixture_matches_example_policy(self) -> None:
        proc = run_parity(FIXTURE)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("PARITY: OK", proc.stdout)

    def test_block_needs_fixture_matches_example_policy(self) -> None:
        # Block-sequence needs, and other jobs declaring needs of their own.
        proc = run_parity(BLOCK_FIXTURE)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("PARITY: OK", proc.stdout)

    def test_every_supported_needs_style_reads_the_same_set(self) -> None:
        styles = {
            "flow": "needs: [lint, unit]\n",
            "flow quoted": "needs: [ 'lint', \"unit\" ]  # comment\n",
            "block": "needs:\n  - lint\n  - unit\n",
            "block at key indent": "needs:\n- lint\n- unit\n",
        }
        for name, needs_yaml in styles.items():
            with self.subTest(style=name):
                proc = run_parity(workflow_with(needs_yaml), policy_for(["lint", "unit"]))
                self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_custom_policy_is_respected(self) -> None:
        # El caller puede apuntar el gate a su propia política: un workflow
        # y una política que acuerdan otro conjunto de jobs no es drift.
        workflow = workflow_with("needs: [lint, unit, e2e]\n",
                                 "e2e:\n  runs-on: ubuntu-latest\n")
        proc = run_parity(workflow, policy_for(["lint", "unit", "e2e"]))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_workflow_job_outside_the_policy_exits_nonzero(self) -> None:
        proc = run_parity(FIXTURE, policy_for(["lint", "compile", "unit", "full"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("PARITY: FAIL", proc.stdout)
        self.assertIn("workflow job 'smoke' is not in the policy", proc.stdout)
        self.assertIn("aggregator needs 'smoke' is not in the policy", proc.stdout)

    def test_policy_job_absent_from_the_workflow_exits_nonzero(self) -> None:
        proc = run_parity(FIXTURE, policy_for(POLICY_JOBS + ["e2e"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("policy job 'e2e' is not defined in the workflow", proc.stdout)

    def test_job_not_wired_into_needs_exits_nonzero(self) -> None:
        # The job exists and the policy knows it, but the aggregator never waits for it.
        proc = run_parity(workflow_with("needs: [lint]\n"), policy_for(["lint", "unit"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("policy job 'unit' is missing from the aggregator needs", proc.stdout)

    def test_needs_entry_outside_the_policy_exits_nonzero(self) -> None:
        proc = run_parity(workflow_with("needs: [lint, unit]\n"), policy_for(["lint"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("aggregator needs 'unit' is not in the policy", proc.stdout)

    def test_job_added_to_the_workflow_only_exits_nonzero(self) -> None:
        extra = "e2e:\n  runs-on: ubuntu-latest\n"
        proc = run_parity(workflow_with("needs: [lint, unit]\n", extra), policy_for(["lint", "unit"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("workflow job 'e2e' is not in the policy", proc.stdout)

    def test_needs_entry_that_is_not_a_workflow_job_exits_nonzero(self) -> None:
        proc = run_parity(workflow_with("needs: [lint, unit, ghost]\n"),
                          policy_for(["lint", "unit", "ghost"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("policy job 'ghost' is not defined in the workflow", proc.stdout)

    def test_scalar_needs_is_read(self) -> None:
        workflow = """\
            jobs:
              lint:
                runs-on: ubuntu-latest
              required:
                needs: lint
                runs-on: ubuntu-latest
            """
        proc = run_parity(workflow, policy_for(["lint"]))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_aggregator_is_chosen_by_id_not_by_having_needs(self) -> None:
        # 'compile' has needs too; naming it as the aggregator must not pass.
        proc = run_parity(BLOCK_FIXTURE, aggregator="compile")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("PARITY: FAIL", proc.stdout)


class UnreadableWorkflowTests(unittest.TestCase):
    """What the parser cannot read is a doubt, named in the message."""

    def assert_fail_closed(self, proc: subprocess.CompletedProcess, fragment: str) -> None:
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("FAIL fail-closed:", proc.stderr)
        self.assertIn(fragment, proc.stderr)
        self.assertNotIn("Traceback", proc.stderr)
        self.assertNotIn("PARITY: OK", proc.stdout)

    def test_unsupported_needs_constructs_exit_nonzero(self) -> None:
        constructs = {
            "expression": "needs: ${{ fromJSON(vars.JOBS) }}\n",
            "alias": "needs: *deps\n",
            "anchor": "needs: &deps [lint, unit]\n",
            "multi-line flow sequence": "needs: [lint,\n  unit]\n",
            "flow mapping": "needs: {a: lint}\n",
            "block scalar": "needs: |\n  lint\n",
            "nested sequence": "needs:\n  - [lint, unit]\n",
            "mapping entry": "needs:\n  - job: lint\n",
            "no entries": "needs: []\n",
            "empty needs": "needs:\n",
            "duplicate needs entry": "needs: [lint, unit, lint]\n",
            "duplicate 'needs' key": "needs: [lint, unit]\nneeds: [lint]\n",
            "merge key": "<<: *defaults\nneeds: [lint, unit]\n",
        }
        for construct, needs_yaml in constructs.items():
            with self.subTest(construct=construct):
                proc = run_parity(workflow_with(needs_yaml), policy_for(["lint", "unit"]))
                self.assert_fail_closed(proc, construct)

    def test_expression_inside_a_block_item_exits_nonzero(self) -> None:
        proc = run_parity(workflow_with("needs:\n  - lint\n  - ${{ matrix.dep }}\n"),
                          policy_for(["lint", "unit"]))
        self.assert_fail_closed(proc, "expression")

    def test_missing_jobs_block_exits_nonzero(self) -> None:
        proc = run_parity("name: ci\non: [push]\n", policy_for(["lint"]))
        self.assert_fail_closed(proc, "no top-level 'jobs:' block")

    def test_missing_aggregator_exits_nonzero(self) -> None:
        proc = run_parity(FIXTURE, aggregator="gate")
        self.assert_fail_closed(proc, "aggregator job 'gate' is not defined")

    def test_aggregator_without_needs_exits_nonzero(self) -> None:
        proc = run_parity(FIXTURE, aggregator="lint")
        self.assert_fail_closed(proc, "aggregator job 'lint' declares no needs")

    def test_job_defined_inline_or_by_alias_exits_nonzero(self) -> None:
        for name, extra in (("flow mapping", "e2e: {runs-on: ubuntu-latest}\n"),
                            ("alias", "e2e: *job\n"),
                            ("anchor", "e2e: &job\n  runs-on: ubuntu-latest\n")):
            with self.subTest(construct=name):
                proc = run_parity(workflow_with("needs: [lint, unit]\n", extra),
                                  policy_for(["lint", "unit"]))
                self.assert_fail_closed(proc, "unsupported job entry")

    def test_duplicate_job_key_exits_nonzero(self) -> None:
        extra = "lint:\n  runs-on: ubuntu-latest\n"
        proc = run_parity(workflow_with("needs: [lint, unit]\n", extra), policy_for(["lint", "unit"]))
        self.assert_fail_closed(proc, "duplicate job 'lint'")

    def test_tab_indentation_exits_nonzero(self) -> None:
        proc = run_parity("jobs:\n\tlint:\n\t\truns-on: x\n", policy_for(["lint"]))
        self.assert_fail_closed(proc, "tab")

    def test_jobs_block_given_inline_exits_nonzero(self) -> None:
        proc = run_parity("jobs: {lint: {runs-on: x}}\n", policy_for(["lint"]))
        self.assert_fail_closed(proc, "'jobs:' with an inline value")

    def test_unreadable_workflow_exits_nonzero(self) -> None:
        proc = run_parity(FIXTURES / "no-such-workflow.yml")
        self.assert_fail_closed(proc, "unreadable workflow file")

    def test_tab_indented_needs_item_exits_nonzero(self) -> None:
        proc = run_parity(workflow_with("needs:\n\t  - lint\n"), policy_for(["lint", "unit"]))
        self.assert_fail_closed(proc, "tab in indentation")

    def test_tab_inside_a_job_body_exits_nonzero(self) -> None:
        workflow = workflow_with("needs: [lint, unit]\n").replace(
            "  lint:\n", "  lint:\n\truns-on: ubuntu-latest\n")
        proc = run_parity(workflow, policy_for(["lint", "unit"]))
        self.assert_fail_closed(proc, "tab in indentation")

    def test_invalid_policy_exits_nonzero(self) -> None:
        proc = run_parity(FIXTURE, {"required_jobs": POLICY_JOBS})
        self.assert_fail_closed(proc, "policy file")

    def test_aggregator_listed_in_the_policy_exits_nonzero(self) -> None:
        proc = run_parity(workflow_with("needs: [lint, unit]\n"),
                          policy_for(["lint", "unit", "required"]))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("policy lists the aggregator 'required' itself", proc.stdout)


class MissingSiblingTests(unittest.TestCase):
    """Installed without check_required_jobs.py: explicit fail-closed, never
    a traceback (#161, native-review correction)."""

    def test_gate_fails_closed_without_the_sibling_gate(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            install = root / "install"
            install.mkdir()
            shutil.copy(SCRIPT, install / "check_parity.py")
            (install / "policy.json").write_text(
                json.dumps({"required_jobs": ["lint"], "events": {"push": {"accepted_skips": []}}}),
                encoding="utf-8")
            (root / "workflow.yml").write_text(
                "jobs:\n  lint:\n    runs-on: ubuntu-latest\n", encoding="utf-8")
            proc = subprocess.run(
                [sys.executable, str(install / "check_parity.py"),
                 "--workflow", str(root / "workflow.yml"),
                 "--policy", str(install / "policy.json"),
                 "--aggregator", "lint"],
                capture_output=True, text=True)
        self.assertNotEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("fail-closed", proc.stderr)
        self.assertIn("check_required_jobs.py", proc.stderr)
        self.assertNotIn("Traceback", proc.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
