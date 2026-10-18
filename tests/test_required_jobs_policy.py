"""The `required` aggregator runs the canon gate (DysTelefonica/team-skills#806, issue #1040).

BP no longer keeps its own `check_required_jobs.py`: the job executes
`.agents/skills/ci-pattern/assets/required-jobs/check_required_jobs.py` against
`.github/required-jobs-policy.json`. These tests run that canon gate, as a
process, against BP's real policy and workflow, so a policy that stops matching
the workflow or the events `ci.yml` listens to fails here before it blocks every PR.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
CANON = ROOT / ".agents" / "skills" / "ci-pattern" / "assets" / "required-jobs"
POLICY = ROOT / ".github" / "required-jobs-policy.json"
WORKFLOW = ROOT / ".github" / "workflows" / "ci.yml"
EVENTS = ("pull_request", "push", "schedule", "workflow_dispatch")


def _policy() -> dict:
    return json.loads(POLICY.read_text(encoding="utf-8"))


def _needs(**overrides: str) -> dict[str, dict[str, str]]:
    needs = {job: {"result": "success"} for job in _policy()["required_jobs"]}
    for job, result in overrides.items():
        needs[job.replace("_", "-")] = {"result": result}
    return needs


def _gate(event: str, needs: dict, policy: Path = POLICY, base: Path | None = None):
    argv = [sys.executable, str(CANON / "check_required_jobs.py"), "--policy", str(policy),
            "--event", event]
    if base is not None:
        argv += ["--base-policy", str(base)]
    return subprocess.run(argv, input=json.dumps(needs), capture_output=True, text=True, check=False)


def test_workflow_needs_and_policy_name_the_same_jobs() -> None:
    result = subprocess.run(
        [sys.executable, str(CANON / "check_parity.py"), "--workflow", str(WORKFLOW),
         "--aggregator", "required", "--policy", str(POLICY)],
        capture_output=True, text=True, check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize("event", EVENTS)
def test_every_event_ci_listens_to_is_declared(event: str) -> None:
    assert event in _policy()["events"], f"ci.yml fires on {event} but the policy does not declare it"
    assert _gate(event, _needs()).returncode == 0


@pytest.mark.parametrize("event", ["push", "schedule", "workflow_dispatch"])
def test_pr_gate_may_be_skipped_outside_pull_requests(event: str) -> None:
    result = _gate(event, _needs(pr_gate="skipped"))
    assert result.returncode == 0, result.stdout + result.stderr


def test_pr_gate_skipped_on_a_pull_request_fails() -> None:
    result = _gate("pull_request", _needs(pr_gate="skipped"))
    assert result.returncode == 1
    assert "pr-gate" in result.stdout + result.stderr


def test_any_other_skip_fails_on_every_event() -> None:
    for event in EVENTS:
        assert _gate(event, _needs(docs_links="skipped")).returncode == 1, event


def test_a_required_job_missing_from_needs_fails() -> None:
    needs = _needs()
    del needs["integration"]
    result = _gate("pull_request", needs)
    assert result.returncode == 1
    assert "integration" in result.stdout + result.stderr


def test_a_needs_key_the_policy_does_not_cover_fails() -> None:
    needs = _needs()
    needs["new-job"] = {"result": "success"}
    assert _gate("pull_request", needs).returncode == 1


def test_a_failed_job_fails_and_leads_the_verdict() -> None:
    result = _gate("pull_request", _needs(quality="failure"))
    assert result.returncode == 1
    assert "quality" in result.stdout + result.stderr


def test_dropping_a_job_from_the_policy_needs_a_relaxation(tmp_path: Path) -> None:
    candidate = _policy()
    candidate["required_jobs"] = [j for j in candidate["required_jobs"] if j != "docs-links"]
    changed = tmp_path / "candidate.json"
    changed.write_text(json.dumps(candidate), encoding="utf-8")
    needs = {job: {"result": "success"} for job in candidate["required_jobs"]}
    result = _gate("pull_request", needs, policy=changed, base=POLICY)
    assert result.returncode == 1
    assert "relaxation" in (result.stdout + result.stderr).lower()
