"""Governed merge of access2web-blueprint (#1359, option A).

The active ruleset `main-maintainers-and-admins-merge` only lets repo admins, maintainers and
the `dys-orquestador` GitHub App update `main`, so a merge made with the workflow's
GITHUB_TOKEN would be rejected. The governed merge therefore mints an installation token of
the App and hands it to the canon `check_governed_merge.py`; it never falls back to
GITHUB_TOKEN. The host contract names that App's bot as the only governed merger, so a merge
made by anybody else is reported by the post-push detective.
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


def test_mints_the_orchestrator_app_token_pinned_by_sha() -> None:
    job = _job()
    mint = _step(job, "actions/create-github-app-token@")
    name, _, ref = mint["uses"].partition("@")
    assert name == "actions/create-github-app-token"
    assert FULL_SHA.fullmatch(ref), mint["uses"]
    assert mint["with"]["app-id"] == "${{ secrets.DYS_ORQUESTADOR_APP_ID }}"
    assert mint["with"]["private-key"] == "${{ secrets.DYS_ORQUESTADOR_PRIVATE_KEY }}"
    assert mint["id"] == "app-token"


def test_missing_secrets_fail_closed_before_minting() -> None:
    job = _job()
    steps = job["steps"]
    guard = _step(job, "DYS_ORQUESTADOR_APP_ID is not set")
    mint = _step(job, "actions/create-github-app-token@")
    assert steps.index(guard) < steps.index(mint)
    assert guard["env"]["APP_ID"] == "${{ secrets.DYS_ORQUESTADOR_APP_ID }}"
    assert guard["env"]["PRIVATE_KEY"] == "${{ secrets.DYS_ORQUESTADOR_PRIVATE_KEY }}"
    assert "exit 2" in guard["run"]


def test_merge_runs_with_the_app_token_never_with_github_token() -> None:
    job = _job()
    merge = _step(job, CANON)
    assert merge["env"]["GH_TOKEN"] == "${{ steps.app-token.outputs.token }}"
    assert merge["env"]["GITHUB_TOKEN"] == "${{ steps.app-token.outputs.token }}"
    assert merge["env"]["PR_ID"] == "${{ inputs.pr }}"
    text = WORKFLOW.read_text(encoding="utf-8")
    assert "github.token" not in text
    assert "secrets.GITHUB_TOKEN" not in text
    # The PR number enters through env, never interpolated into a script.
    assert "${{ inputs.pr }}" not in merge["run"]
    assert '--pr "$PR_ID"' in merge["run"]


def test_workflow_token_stays_read_only() -> None:
    assert _workflow()["permissions"] == {"contents": "read"}
    assert "permissions" not in _job()


def test_contract_names_the_orchestrator_bot_as_the_only_governed_merger() -> None:
    contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
    assert contract["governed_path"] == {
        "enabled": True,
        "actors": ["dys-orquestador[bot]"],
    }
    assert [c["name"] for c in contract["required_checks"]] == ["required", "pr-gate"]
