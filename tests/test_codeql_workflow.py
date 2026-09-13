from pathlib import Path

import yaml

WORKFLOW = Path(__file__).resolve().parent.parent / ".github" / "workflows" / "codeql.yml"
CODEQL_ACTION_SHA = "cdf488f595d80d6e07e03d4674febd5ab45fa938"


def _load_workflow() -> dict:
    return yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))


def _triggers(workflow: dict) -> dict:
    # PyYAML 1.1 resolves the unquoted GitHub Actions key `on` as boolean true.
    return workflow.get("on", workflow.get(True, {}))


def test_codeql_scans_python_via_reusable_workflow_call() -> None:
    """Issue #702: codeql.yml is `on: workflow_call` only.

    It no longer has independent `pull_request`/`push`/`schedule`/
    `workflow_dispatch` triggers of its own — it is invoked as a job
    (`uses: ./.github/workflows/codeql.yml`) from `ci.yml`, which is what
    lets the `required` aggregator `needs:` it (`needs:` cannot cross
    workflow files, but a job invoked via `uses:` still reports an
    ordinary job result). `ci.yml`'s own triggers (pull_request, push to
    main, weekly schedule, workflow_dispatch) are what now drive when this
    analysis actually runs.
    """
    workflow = _load_workflow()
    triggers = _triggers(workflow)
    job = workflow["jobs"]["codeql"]

    assert set(triggers) == {"workflow_call"}
    assert job["runs-on"] == "ubuntu-24.04"
    assert job["permissions"]["security-events"] == "write"

    init = next(step for step in job["steps"] if step.get("name") == "Initialize CodeQL")
    analyze = next(step for step in job["steps"] if step.get("name") == "Analyze Python")
    assert init["uses"] == f"github/codeql-action/init@{CODEQL_ACTION_SHA}"
    assert init["with"] == {"languages": "python", "build-mode": "none"}
    assert analyze["uses"] == f"github/codeql-action/analyze@{CODEQL_ACTION_SHA}"
