"""Post-push detective of access2web-blueprint (#1359).

The host contract declares `governed_path` (enabled, actors ["dys-orquestador[bot]"]), but a
declaration nobody measures is decoration. This workflow runs the vendored canon
`check_push_compliance.py` after every push to `main`: each pushed commit must come from a
merged PR whose required checks (`required`, `pr-gate`) are green on its head, and that PR must
have been merged by a governed actor (or carry the governed merge marker). A violation opens an
incident issue and fails the job (exit 1); an unmeasurable state fails closed (exit 2).

The governed merge merges with the dys-orquestador App token, so its merges trigger a normal
`push` event: no `post_merge_dispatch` is needed for the detective to see them.
"""

import json
import re
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
WORKFLOW = ROOT / ".github" / "workflows" / "push-compliance.yml"
CONTRACT = ROOT / ".github" / "host-contract.json"
ADOPTION = ROOT / ".github" / "ci-pattern-adoption.json"
FULL_SHA = re.compile(r"[0-9a-f]{40}")
CANON = ".agents/skills/ci-pattern/assets/runner-controls/check_push_compliance.py"


def _workflow() -> dict:
    return yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))


def _triggers(workflow: dict) -> dict:
    # PyYAML 1.1 resolves the unquoted GitHub Actions key `on` as boolean true.
    return workflow.get("on", workflow.get(True, {}))


def _job() -> dict:
    return _workflow()["jobs"]["push-compliance"]


def _step(job: dict, needle: str) -> dict:
    matches = [s for s in job["steps"] if needle in str(s.get("uses", "")) + str(s.get("run", ""))]
    assert len(matches) == 1, f"expected exactly one step with {needle!r}, got {len(matches)}"
    return matches[0]


def test_triggers_on_push_to_main_only() -> None:
    triggers = _triggers(_workflow())
    assert set(triggers) == {"push"}
    assert triggers["push"] == {"branches": ["main"]}


def test_runs_the_vendored_canon_against_the_host_contract() -> None:
    audit = _step(_job(), CANON)
    run = audit["run"]
    assert f"python3 {CANON}" in run
    assert '--event "$GITHUB_EVENT_PATH"' in run
    assert "--contract .github/host-contract.json" in run
    # The canon imports sibling modules: `python3 -I` would drop its directory from sys.path.
    assert "python3 -I" not in run


def test_uses_the_read_scoped_workflow_token_and_a_hosted_runner() -> None:
    workflow = _workflow()
    assert workflow["permissions"] == {
        "contents": "read",
        "checks": "read",
        "issues": "write",
        "pull-requests": "read",
    }
    job = _job()
    assert "permissions" not in job
    assert job["runs-on"] == "ubuntu-24.04"
    assert job["env"]["GH_TOKEN"] == "${{ github.token }}"
    # The canon waits up to 900 s for the required checks; the job must outlast that ceiling.
    assert job["timeout-minutes"] > 15
    checkout = _step(job, "actions/checkout@")
    assert FULL_SHA.fullmatch(checkout["uses"].partition("@")[2]), checkout["uses"]
    assert checkout["with"]["persist-credentials"] is False


def test_audits_are_serialized_without_cancelling() -> None:
    concurrency = _workflow()["concurrency"]
    assert concurrency["cancel-in-progress"] is False
    assert "${{ github.ref }}" in concurrency["group"]


def test_contract_measured_by_the_detective_names_the_orchestrator_bot() -> None:
    contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
    assert contract["protected_branches"] == ["main"]
    assert contract["governed_path"] == {
        "enabled": True,
        "actors": ["dys-orquestador[bot]"],
    }
    assert [c["name"] for c in contract["required_checks"]] == ["required", "pr-gate"]


def test_workflow_is_inventoried_as_adopted() -> None:
    adoption = json.loads(ADOPTION.read_text(encoding="utf-8"))
    targets = {
        entry.get("target")
        for entry in adoption["phases"]["0"]["inventory"]
        if entry.get("disposition") == "adopted"
    }
    assert ".github/workflows/push-compliance.yml" in targets
