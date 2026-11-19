"""Governed merge of access2web-blueprint, without Apps or roles (owner's decision, 2026-11-19).

The administrator running the session dispatches the workflow with their own `gh`; the canon
`check_governed_merge.py` verifies the PR contract and merges with the workflow's GITHUB_TOKEN,
as in DysTelefonica/cadete. The ruleset on main requires a PR and the required checks
(`required`, `pr-gate`) in green, so the token can only merge a green PR. A merge made with
GITHUB_TOKEN triggers no `push`, so the contract dispatches the post-push detective afterwards
(`post_merge_dispatch`, HR-817).
"""

import json
import re
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
WORKFLOW = ROOT / ".github" / "workflows" / "governed-merge.yml"
CONTRACT = ROOT / ".github" / "host-contract.json"
FULL_SHA = re.compile(r"[0-9a-f]{40}")
CANON = ".agents/skills/ci-pattern/assets/runner-controls/check_governed_merge.py"


def _workflow() -> dict:
    return yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))


def _triggers(workflow: dict) -> dict:
    # PyYAML 1.1 resolves the unquoted GitHub Actions key `on` as boolean true.
    return workflow.get("on", workflow.get(True, {}))


def _job() -> dict:
    return _workflow()["jobs"]["governed-merge"]


def _step(job: dict, needle: str) -> dict:
    matches = [s for s in job["steps"] if needle in str(s.get("uses", "")) + str(s.get("run", ""))]
    assert len(matches) == 1, f"expected exactly one step with {needle!r}, got {len(matches)}"
    return matches[0]


def test_dispatch_only_with_pr_input_and_only_from_main() -> None:
    triggers = _triggers(_workflow())
    assert set(triggers) == {"workflow_dispatch"}
    assert triggers["workflow_dispatch"]["inputs"]["pr"]["required"] is True
    assert _job()["if"] == "github.ref == 'refs/heads/main'"


def test_no_app_token_is_minted() -> None:
    text = WORKFLOW.read_text(encoding="utf-8")
    assert "create-github-app-token" not in text
    assert "secrets.DYS_" not in text


def test_merge_runs_with_the_workflow_token() -> None:
    job = _job()
    merge = _step(job, CANON)
    assert merge["env"]["GH_TOKEN"] == "${{ github.token }}"
    assert "GITHUB_TOKEN" not in merge["env"]
    assert merge["env"]["PR_ID"] == "${{ inputs.pr }}"
    # The PR number enters through env, never interpolated into a script.
    assert "${{ inputs.pr }}" not in merge["run"]
    assert '--pr "$PR_ID"' in merge["run"]


def test_only_the_job_gets_write_permissions() -> None:
    assert _workflow()["permissions"] == {"contents": "read"}
    assert _job()["permissions"] == {
        "contents": "write",
        "issues": "write",
        "checks": "read",
        "pull-requests": "read",
        "actions": "write",
    }


def test_contract_names_the_workflow_token_as_the_only_governed_merger() -> None:
    contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
    assert contract["governed_path"] == {
        "enabled": True,
        "actors": ["github-actions[bot]"],
        "merge_marker": r"^Merge PR #\d+ \((safe|governed) merge\)$",
    }
    assert [c["name"] for c in contract["required_checks"]] == ["required", "pr-gate"]
