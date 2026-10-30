"""#1464: the `required` aggregator must not publish the host-required check name
outside the events that evaluate the pull request (HR-30).

The host contract of this repository declares the required check `required`
(`.github/host-contract.json`: `required_checks[0].name = "required"`, class
`host-enforced`), alongside the H1 context `pr-gate`. The aggregator
published that name from every event the workflow
listens to, so a manual `workflow_dispatch` over a pull request branch published a
green `required` even though `pr-gate` — the gate of the issue link and the review
budget — had never run.

HR-30 is explicit: a required check name must be published only from an event that
evaluates the pull request; in any other event the job must publish under another
context name or fail, and it must not emit a success it did not earn.

The same class of defect is invisible while the production gate measures the
workflows without the host contract: the canon's `workflow-policy` gate only knows
which names the host requires when it is given `--host-contract`. The production
invocation did not pass it, so the defect could not be detected in CI.

These checks take the workflow and the policy as they are on disk, and the canon's
own `allowed_events` decides which events a job may run on, so the assertions are
the gate's semantics and not a reading of the `if:` text.
"""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
CI_WORKFLOW = ROOT / ".github" / "workflows" / "ci.yml"
HOST_CONTRACT = ROOT / ".github" / "host-contract.json"
POLICY = ROOT / ".github" / "required-jobs-policy.json"
CANON_GATE = (
    ROOT
    / ".agents"
    / "skills"
    / "ci-pattern"
    / "assets"
    / "workflow-policy"
    / "check_workflow_policy.py"
)

#: The events the aggregator may run on: the ones that evaluate the pull request
#: and the push to the base branch (a push run reports on the base, not on a pull
#: request head, so it satisfies the branch protection of `main`).
PR_EVENTS = {"pull_request", "push"}

#: The events that never evaluate a pull request.
NON_PR_EVENTS = {"schedule", "workflow_dispatch"}


def _canon_gate():
    """The canon's workflow-policy gate, loaded from the vendored catalog."""
    spec = importlib.util.spec_from_file_location("canon_workflow_policy", CANON_GATE)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _workflow() -> dict:
    return yaml.safe_load(CI_WORKFLOW.read_text(encoding="utf-8"))


def _policy() -> dict:
    return json.loads(POLICY.read_text(encoding="utf-8"))


def test_the_required_aggregator_runs_only_on_the_events_that_evaluate_the_pull_request() -> None:
    """The aggregator publishes the host-required name only from `pull_request` and `push`.

    A scheduled run or a manual dispatch does not evaluate a pull request, so a green
    `required` from there is a false green: it is indistinguishable from the real one
    and it lets a pull request through without `pr-gate` ever running.
    """
    module = _canon_gate()
    document = _workflow()
    job = document["jobs"]["required"]
    events = module.workflow_events(document)
    allowed = module.allowed_events(job, events)
    assert allowed is not None, "the aggregator's `if:` must be provable by the canon's gate"
    assert not (allowed & NON_PR_EVENTS), (
        f"the aggregator publishes the host-required name on {sorted(allowed & NON_PR_EVENTS)}: "
        "HR-30 requires it only from an event that evaluates the pull request"
    )
    assert PR_EVENTS <= allowed, (
        f"the aggregator must keep running on {sorted(PR_EVENTS)}: a pull request needs the "
        "verdict published under the name the host contract requires"
    )


def test_the_production_gate_measures_the_workflows_against_the_host_contract() -> None:
    """The production invocation passes `--host-contract`.

    Without it the gate cannot know which check names the host requires, so it cannot
    apply HR-30 to them and the defect stays invisible in CI.
    """
    document = _workflow()
    quality = document["jobs"]["quality"]
    invocations = [
        step["run"]
        for step in quality["steps"]
        if "check_workflow_policy.py" in step.get("run", "")
    ]
    assert len(invocations) == 1, "the quality job must run the canon's workflow-policy gate once"
    assert "--host-contract .github/host-contract.json" in invocations[0], (
        "the production gate must measure against the host contract, or HR-30 has nothing to check"
    )


def test_the_workflow_policy_gate_reports_no_finding_against_the_host_contract() -> None:
    """The canon's gate, run the way production runs it, is green.

    This is the assertion that pins the defect: the same command with the host
    contract used to report two HR-30 findings for the `required` job.
    """
    workflows = sorted((ROOT / ".github" / "workflows").glob("*.yml"))
    command = [
        sys.executable,
        str(CANON_GATE),
        # Relative paths, like the production invocation: the gate compares the
        # paths it is given with the ones the policy declares, and the policy
        # declares `.github/workflows/ci.yml`.
        *[
            argument
            for path in workflows
            for argument in ("--workflow", str(path.relative_to(ROOT)))
        ],
        "--base-branch",
        "main",
        "--required-jobs",
        str(POLICY.relative_to(ROOT)),
        "--host-contract",
        str(HOST_CONTRACT.relative_to(ROOT)),
    ]
    result = subprocess.run(command, capture_output=True, text=True, check=False, cwd=ROOT)
    assert result.returncode == 0, (
        "the canon's workflow-policy gate reports findings against the host contract:\n"
        f"{result.stdout}{result.stderr}"
    )


def test_every_accepted_skip_of_the_nightly_events_is_justified() -> None:
    """Each accepted skip of `schedule` and `workflow_dispatch` is named, with its event.

    The relaxation is the contract of what the policy accepts. A skip that no
    relaxation explains is an invisible hole: the aggregator accepts it and the pull
    request goes green without the job that would have evaluated it.
    """
    policy = _policy()
    relaxation = policy["relaxation"]
    reason = relaxation["reason"]
    assert relaxation.get("reference") and relaxation.get("issue"), (
        "the relaxation must carry a reference and an issue that identify it"
    )
    for event in sorted(NON_PR_EVENTS):
        for skip in policy["events"][event]["accepted_skips"]:
            assert skip in reason, (
                f"the accepted skip '{skip}' of the event '{event}' is not justified in the "
                "relaxation: name the skip and the event that accepts it"
            )
        assert event in reason, (
            f"the relaxation must name the event '{event}' whose skips it justifies"
        )
