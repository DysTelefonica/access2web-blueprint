"""Wiring of the H1 gate of the pull request (BP #1461).

The model (canon #1270): the workflow runs in `pull_request_target`, so what
runs is the code of the BASE. The job `canon` resolves the catalog at the seal
declared in `.team-skills.yaml` and publishes it as an artifact; the gate job
consumes that artifact, holds no secret, reads the head of the pull request as
data and publishes the verdict as a check run on that head.

These tests are WIRING: they prove that the YAML says what the model requires
(what runs, with which token, from which tree, under which name). They do not
replace the behaviour test of the canon's own gates, which the catalog runs
against its own fixtures.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest
import yaml

ROOT = Path(__file__).resolve().parent.parent
WORKFLOWS = ROOT / ".github" / "workflows"
H1_WORKFLOW = WORKFLOWS / "pr-gates.yml"
LEGACY_WORKFLOW = WORKFLOWS / "pr-gate.yml"
ACTION = ROOT / ".github" / "actions" / "pattern-at-pin" / "action.yml"

#: The context the host contract requires: the name of the workflow and the
#: name of the gate job. With another name the required check does not appear
#: on the head of the pull request and the gate does not block the merge.
H1_CONTEXT = "pr-gate"

#: The events that can change the verdict of the gates.
H1_EVENTS = ("opened", "edited", "labeled", "unlabeled", "synchronize", "reopened")

#: The gates that the legacy workflow ran against the tree of the pull request.
PR_GATES = ("pr size", "issue spec")


@pytest.fixture(scope="module")
def h1() -> dict:
    """The workflow that owns the H1 gate."""
    assert H1_WORKFLOW.is_file(), f"{H1_WORKFLOW} does not exist: the H1 gate lives there"
    return yaml.safe_load(H1_WORKFLOW.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def h1_text() -> str:
    return H1_WORKFLOW.read_text(encoding="utf-8")


@pytest.fixture(scope="module")
def legacy() -> dict:
    """The workflow of `pull_request` that keeps the context and the lanes."""
    return yaml.safe_load(LEGACY_WORKFLOW.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def action() -> dict:
    assert ACTION.is_file(), f"{ACTION} does not exist: the canon is resolved by that action"
    return yaml.safe_load(ACTION.read_text(encoding="utf-8"))


def _triggers(doc: dict) -> dict:
    """The `on:` block; PyYAML parses the plain key `on` as the boolean `True`."""
    return doc.get(True) or doc.get("on") or {}


def _steps(job: dict) -> list[dict]:
    return job.get("steps") or []


def _run_blocks(job: dict) -> list[str]:
    return [step["run"] for step in _steps(job) if "run" in step]


def _checkouts(job: dict) -> list[dict]:
    return [
        step for step in _steps(job) if str(step.get("uses", "")).startswith("actions/checkout@")
    ]


# --- The shape of the H1 workflow -------------------------------------------------------------


def test_the_h1_gate_runs_on_pull_request_target(h1: dict) -> None:
    """In `pull_request` the workflow that runs is the one of the head.

    The token of the catalog cannot be in a job that executes the code of the
    pull request: `pull_request_target` runs the workflow of the base.
    """
    triggers = _triggers(h1)
    assert "pull_request_target" in triggers, (
        f"the H1 workflow must run in `pull_request_target`, got {sorted(triggers)}"
    )
    assert "pull_request" not in triggers, (
        "`pull_request` runs the workflow of the head: the catalog token would "
        "reach the code of the pull request"
    )
    assert list(triggers["pull_request_target"]["types"]) == list(H1_EVENTS), (
        "the events that change the body or the labels of the pull request must "
        f"re-run the gate: {list(H1_EVENTS)}"
    )


def test_the_h1_workflow_has_a_canon_job_and_a_gate_job(h1: dict) -> None:
    """The gate consumes what the `canon` job published."""
    jobs = h1["jobs"]
    assert set(jobs) == {"canon", "pr-gate"}, f"unexpected jobs: {sorted(jobs)}"
    assert jobs["pr-gate"]["needs"] == "canon", (
        "the gate runs the canon published by the job `canon`, not a copy of it"
    )


def test_the_canon_job_checks_out_the_base_without_credentials(h1: dict) -> None:
    """The action and the seal come from the base.

    The pull request can neither modify the action nor choose the commit that
    measures it.
    """
    checkouts = _checkouts(h1["jobs"]["canon"])
    assert len(checkouts) == 1, "the job `canon` checks out the base exactly once"
    checkout = checkouts[0]
    assert checkout["with"]["ref"] == "${{ github.event.pull_request.base.sha || github.sha }}"
    assert checkout["with"]["persist-credentials"] is False, (
        "a checkout that persists the token leaves it in the git config of the job"
    )


def test_the_catalog_token_lives_only_in_the_canon_step(h1: dict, h1_text: str) -> None:
    """The token of the catalog appears in one step of one job."""
    assert h1_text.count("TEAM_SKILLS_READ_TOKEN") == 1, (
        "the token of the catalog must appear in a single step of the workflow"
    )
    canon = yaml.safe_dump(h1["jobs"]["canon"])
    assert "TEAM_SKILLS_READ_TOKEN" in canon, "the token belongs to the job `canon`"
    gate = yaml.safe_dump(h1["jobs"]["pr-gate"])
    assert "TEAM_SKILLS_READ_TOKEN" not in gate, "the gate job does not read the catalog"
    action_steps = [
        step
        for step in _steps(h1["jobs"]["canon"])
        if str(step.get("uses", "")).endswith("pattern-at-pin")
    ]
    assert len(action_steps) == 1, "the action that resolves the canon runs exactly once"
    assert action_steps[0]["with"]["token"] == "${{ secrets.TEAM_SKILLS_READ_TOKEN }}"


def test_the_gate_job_holds_no_secret(h1: dict) -> None:
    """The job that measures the pull request holds no secret at all."""
    gate = yaml.safe_dump(h1["jobs"]["pr-gate"])
    assert "secrets." not in gate, "the gate job must not reference any secret"


def test_the_gate_job_checks_out_the_base_and_the_head_in_separate_paths(h1: dict) -> None:
    """The head of the pull request enters as data, in its own path."""
    checkouts = _checkouts(h1["jobs"]["pr-gate"])
    assert len(checkouts) == 2, "the gate job checks out the base and the head"
    paths = {checkout["with"]["path"] for checkout in checkouts}
    assert paths == {"base", "pr-head"}, f"unexpected checkout paths: {sorted(paths)}"
    for checkout in checkouts:
        assert checkout["with"]["persist-credentials"] is False, (
            "every checkout of a pull request job must leave no credentials behind"
        )
    refs = {checkout["with"]["path"]: checkout["with"]["ref"] for checkout in checkouts}
    assert refs["pr-head"] == "${{ github.event.pull_request.head.sha || github.sha }}"
    assert refs["base"] == "${{ github.event.pull_request.base.sha }}"


def test_the_canon_travels_to_the_gate_as_an_artifact(h1: dict) -> None:
    """The canon is published by `canon` and unpacked by the gate."""
    upload = [
        step
        for step in _steps(h1["jobs"]["canon"])
        if str(step.get("uses", "")).startswith("actions/upload-artifact@")
    ]
    assert len(upload) == 1, "the job `canon` publishes the canon once"
    assert upload[0]["with"]["name"] == "canon"

    download = [
        step
        for step in _steps(h1["jobs"]["pr-gate"])
        if str(step.get("uses", "")).startswith("actions/download-artifact@")
    ]
    assert len(download) == 1, "the gate job downloads the canon once"
    assert download[0]["with"]["name"] == "canon"
    assert download[0]["with"]["path"] == ".canon"

    unpacked = "\n".join(_run_blocks(h1["jobs"]["pr-gate"]))
    assert "tar -xzf .canon/canon.tar.gz -C .canon" in unpacked, (
        "the canon arrives as a tarball and is unpacked in `.canon`"
    )


def test_the_gates_run_from_the_published_canon(h1: dict) -> None:
    """The scripts of the gate are the ones of the canon, not of the pull request."""
    assert h1["jobs"]["pr-gate"]["env"]["PATTERN_DIR"] == ".canon"
    blocks = "\n".join(_run_blocks(h1["jobs"]["pr-gate"]))
    assert "$PATTERN_DIR/assets/pr-size/check_pr_size.py" in blocks
    assert "$PATTERN_DIR/assets/issue-spec/check_issue_spec.py" in blocks
    assert "pr-head/assets" not in blocks, (
        "the gate must not execute the scripts that come with the pull request"
    )


def test_the_verdict_is_published_as_a_check_run_on_the_head(h1: dict) -> None:
    """A run of `pull_request_target` is attributed to the base.

    Its verdict only blocks the merge when it is published as a check run on
    the head of the pull request, under the name the host contract requires.
    """
    job = h1["jobs"]["pr-gate"]
    assert job["permissions"]["checks"] == "write", "publishing a check run needs `checks: write`"

    publish = [step for step in _steps(job) if "check-runs" in str(step.get("run", ""))]
    assert len(publish) == 1, "the gate job publishes its verdict once"
    step = publish[0]
    assert step["env"]["CONTEXT"] == H1_CONTEXT, (
        f"the name of the check run must be exactly {H1_CONTEXT!r}"
    )
    assert step["env"]["HEAD_SHA"] == "${{ github.event.pull_request.head.sha }}"
    assert "POST" in step["run"] and "check-runs" in step["run"]
    assert "head_sha" in step["run"], "the check run is bound to the head of the pull request"


def test_every_action_is_pinned_by_sha_and_the_runners_are_hosted(h1: dict, legacy: dict) -> None:
    """The actions run a known commit and the jobs run on hosted runners."""
    for name, doc in (("pr-gates.yml", h1), ("pr-gate.yml", legacy)):
        for job_id, job in doc["jobs"].items():
            assert job.get("runs-on") == "ubuntu-24.04", (
                f"{name}:{job_id} must run on a GitHub-hosted runner"
            )
            for step in _steps(job):
                uses = str(step.get("uses", ""))
                if not uses.startswith("actions/"):
                    continue
                pin = uses.rsplit("@", 1)[1]
                assert re.fullmatch(r"[0-9a-f]{40}", pin), (
                    f"{name}:{job_id} uses `{uses}`, which is not pinned to a commit SHA"
                )


# --- What the legacy workflow keeps and what it stops doing -------------------------------------


def test_the_legacy_workflow_no_longer_runs_the_pr_gates(legacy: dict) -> None:
    """The gates of the pull request moved to the H1 workflow.

    Running them in `pull_request` meant executing the scripts of the tree of
    the pull request with the token of the runner.
    """
    job = legacy["jobs"]["pr-gate"]
    names = [step.get("name") for step in _steps(job)]
    for gate in PR_GATES:
        assert gate not in names, (
            f"the step `{gate}` must not run in `pull_request` anymore: the H1 workflow runs it"
        )
    blocks = "\n".join(_run_blocks(job))
    assert "check_pr_size.py" not in blocks
    assert "check_issue_spec.py" not in blocks


def test_the_legacy_job_publishes_its_own_context(legacy: dict) -> None:
    """Its check run cannot be the context of the H1 gate.

    The legacy job publishes `pr-context`; the H1 gate publishes `pr-gate`.
    Without a name of its own, the legacy job would collide with the H1
    verdict on the head of the pull request.
    """
    assert legacy["jobs"]["pr-gate"]["name"] == "pr-context"


def test_the_legacy_job_holds_no_token_and_no_issue_permission(legacy: dict) -> None:
    """What is left of the legacy job does not need the token of the runner."""
    job = legacy["jobs"]["pr-gate"]
    assert job["permissions"] == {"contents": "read"}, (
        "the gates that needed `issues` and `pull-requests` moved to the H1 workflow"
    )
    assert "GITHUB_TOKEN" not in yaml.safe_dump(job), "the legacy job does not read the token"


# --- The action that resolves the canon --------------------------------------------------------


def test_the_action_is_composite_and_resolves_the_seal(action: dict) -> None:
    """The seal of `.team-skills.yaml` is the only source of the commit."""
    assert action["runs"]["using"] == "composite", "the action is composite"
    blocks = "\n".join(str(step.get("run", "")) for step in action["runs"]["steps"])
    assert ".team-skills.yaml" in blocks, "the seal is read from `.team-skills.yaml`"
    assert "catalog @ " in blocks, "the seal declares the commit as `catalog @ <sha>`"


def test_the_action_clones_the_catalog_with_the_token_in_one_step(action: dict) -> None:
    """The token of the catalog travels in a single step."""
    with_token = [
        step
        for step in action["runs"]["steps"]
        if "TEAM_SKILLS_READ_TOKEN" in yaml.safe_dump(step)
        or "inputs.token" in yaml.safe_dump(step)
    ]
    assert len(with_token) == 1, (
        "the token of the catalog must appear in exactly one step of the action"
    )
    block = str(with_token[0].get("run", ""))
    assert "clone" in block, "that step is the one that clones the catalog"
    assert "DysTelefonica/team-skills" in yaml.safe_dump(action), (
        "the catalog is DysTelefonica/team-skills"
    )


def test_the_action_exports_the_pattern_dir_and_fails_closed(action: dict) -> None:
    """The path of the canon travels in `PATTERN_DIR` and every hole fails loud."""
    blocks = "\n".join(str(step.get("run", "")) for step in action["runs"]["steps"])
    assert "PATTERN_DIR=" in blocks and '>> "$GITHUB_ENV"' in blocks, (
        "the action publishes the directory of the canon in `PATTERN_DIR`"
    )
    assert "personal/ardelperal/ci-pattern" in blocks, (
        "`PATTERN_DIR` points to the ci-pattern skill, not to the root of the catalog"
    )
    assert "mktemp -d" in blocks, "the scratch directory is created with `mktemp -d`"
    assert "exit 1" in blocks, "the action fails closed when it cannot measure"
    assert "checkout" in blocks, "the action checks out the commit of the seal"


def test_the_action_never_deletes_a_directory(action: dict) -> None:
    """The scratch directory is not removed: nothing deletes what it did not create."""
    text = ACTION.read_text(encoding="utf-8")
    assert "rm -" not in text, "the action does not delete any directory"


def test_the_host_contract_requires_the_h1_context() -> None:
    """The verdict blocks the merge only if the host requires its exact context."""
    contract = json.loads((ROOT / ".github" / "host-contract.json").read_text(encoding="utf-8"))
    assert {"name": H1_CONTEXT, "class": "host-enforced"} in contract["required_checks"]

    snapshot = json.loads(
        (ROOT / "docs" / "adoption" / "branch-protection.json").read_text(encoding="utf-8")
    )
    contexts = snapshot["required_status_checks"]["contexts"]
    assert H1_CONTEXT in contexts
    assert "required" in contexts
