from pathlib import Path

import yaml

WORKFLOW = Path(__file__).resolve().parent.parent / ".github" / "workflows" / "codeql.yml"
CODEQL_ACTION_SHA = "cdf488f595d80d6e07e03d4674febd5ab45fa938"


def _load_workflow() -> dict:
    return yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))


def _triggers(workflow: dict) -> dict:
    # PyYAML 1.1 resolves the unquoted GitHub Actions key `on` as boolean true.
    return workflow.get("on", workflow.get(True, {}))


def test_codeql_scans_python_on_pr_push_and_schedule() -> None:
    workflow = _load_workflow()
    triggers = _triggers(workflow)
    job = workflow["jobs"]["codeql"]

    assert {"pull_request", "push", "schedule", "workflow_dispatch"} <= set(triggers)
    assert triggers["pull_request"]["branches"] == ["main"]
    assert triggers["push"]["branches"] == ["main"]
    assert job["runs-on"] == "ubuntu-24.04"
    assert job["permissions"]["security-events"] == "write"

    init = next(step for step in job["steps"] if step.get("name") == "Initialize CodeQL")
    analyze = next(step for step in job["steps"] if step.get("name") == "Analyze Python")
    assert init["uses"] == f"github/codeql-action/init@{CODEQL_ACTION_SHA}"
    assert init["with"] == {"languages": "python", "build-mode": "none"}
    assert analyze["uses"] == f"github/codeql-action/analyze@{CODEQL_ACTION_SHA}"
