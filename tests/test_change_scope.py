"""The surface lanes of the heavy jobs (issue #1376).

``e2e-smoke`` and ``codeql`` only look at the code they analyze, so the
``pr-gate`` publishes one marker per surface and each job reads it. The
markers are declared in ``.github/required-jobs-policy.json`` as
``marker_lanes``: the ``required`` aggregator accepts the skip of a job only
when the marker of the lane is exactly ``"false"``.

The pins below are written against the CONTRACT, not against the shape of the
workflow: the names of the markers live in the policy, and the tests check
that the workflow, the gate and the policy agree. A pin that hard-coded the
``if:`` of a job would fail every time the workflow changed, which is the
churn issue #1375 retired.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
import yaml

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "scripts" / "check_change_scope.py"
POLICY = ROOT / ".github" / "required-jobs-policy.json"
WORKFLOW = ROOT / ".github" / "workflows" / "ci.yml"
PR_GATE = ROOT / ".github" / "workflows" / "pr-gate.yml"

#: The marker of each surface and the job it gates.
SURFACE_JOBS = {
    "e2e_scope_changed": "e2e-smoke",
    "codeql_scope_changed": "codeql",
    "lockfile_scope_changed": "pip-audit",
}


def _classify(surface: str, *paths: str) -> str:
    """Run the classifier over ``paths`` (one per line) and return its verdict."""
    result = subprocess.run(
        [sys.executable, str(SCRIPT), surface],
        input="".join(f"{path}\n" for path in paths),
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    return result.stdout.strip()


def _policy() -> dict:
    return json.loads(POLICY.read_text(encoding="utf-8"))


def _workflow() -> dict:
    return yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))


def _pr_gate() -> dict:
    return yaml.safe_load(PR_GATE.read_text(encoding="utf-8"))


def _lanes() -> list[dict]:
    return _policy().get("marker_lanes", [])


# --- The classifier -------------------------------------------------------


@pytest.mark.parametrize(
    ("surface", "paths", "expected"),
    [
        ("e2e", ("app/Dockerfile", "app/migrations/env.py", "e2e/smoke.spec.ts"), "true"),
        ("codeql", ("app/src/modules/lanzadera/cli.py", "scripts/check_workflows.py"), "true"),
        ("e2e", ("docs/codebase/testing-and-ci.md", "README.md"), "false"),
        ("codeql", (".github/workflows/ci.yml",), "false"),
        # A sibling directory that only looks like the surface is not the surface.
        ("e2e", ("app-legacy/old.py", "scripts-archive/old.py", "appx/x.py"), "false"),
        # Nothing changed, nothing to run.
        ("e2e", (), "false"),
    ],
)
def test_the_classifier_answers_the_surface(
    surface: str, paths: tuple[str, ...], expected: str
) -> None:
    """Each job analyzes one surface; a diff outside it does not pay for the job."""
    assert _classify(surface, *paths) == expected


def test_an_unknown_surface_is_refused() -> None:
    """A typo in the surface must fail loudly, not answer `false` and skip the job."""
    result = subprocess.run(
        [sys.executable, str(SCRIPT), "e2ee"],
        input="app/src/x.py\n",
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 2
    assert "usage" in result.stderr


def test_the_lockfile_surface_covers_the_lock_files() -> None:
    """The dependency audit reads the hash-locked set; only its lock files are its surface."""
    assert _classify("lockfile", "app/requirements-dev.lock") == "true"
    assert _classify("lockfile", "app/uv.lock") == "true"
    assert _classify("lockfile", "app/pyproject.toml") == "false"
    assert _classify("lockfile", "app/src/modules/lanzadera/cli.py") == "false"


def test_a_rename_out_of_the_surface_keeps_the_markers_true(tmp_path: Path) -> None:
    """A rename touches BOTH paths, and the diff must say so.

    `git diff --name-only` reports only the destination of a rename, so moving
    `app/moved.py` to `app-legacy/moved.py` produced a one-path list: the app
    surface looked untouched and the heavy jobs skipped a pull request that had
    just removed application code. The step must diff with `--no-renames`.
    """
    repo = tmp_path / "repo"
    (repo / "scripts").mkdir(parents=True)
    (repo / "app").mkdir()
    (repo / "app" / "moved.py").write_text("x = 1\n", encoding="utf-8")
    for name in ("check_change_scope.py", "check_docs_only_lane.py"):
        shutil.copy2(ROOT / "scripts" / name, repo / "scripts" / name)

    def git(*args: str) -> str:
        return subprocess.run(
            ["git", *args], cwd=repo, check=True, capture_output=True, text=True
        ).stdout.strip()

    git("init", "-q")
    git("config", "user.email", "probe@example.invalid")
    git("config", "user.name", "Probe")
    git("add", ".")
    git("commit", "-qm", "base")
    git("update-ref", "refs/remotes/origin/main", git("rev-parse", "HEAD"))
    (repo / "app-legacy").mkdir()
    git("mv", "app/moved.py", "app-legacy/moved.py")
    git("commit", "-qm", "rename")

    step = next(
        s
        for s in yaml.safe_load(PR_GATE.read_text(encoding="utf-8"))["jobs"]["pr-gate"]["steps"]
        if s.get("id") == "diff-lanes"
    )
    output = tmp_path / "github-output"
    output.write_text("", encoding="utf-8")
    subprocess.run(
        ["bash", "-c", step["run"]],
        cwd=repo,
        env={**os.environ, "BASE_REF": "main", "GITHUB_OUTPUT": str(output)},
        check=True,
    )
    values = dict(
        line.split("=", 1) for line in output.read_text(encoding="utf-8").splitlines() if line
    )
    assert values["docs_changed"] == "false"
    assert values["e2e_scope_changed"] == "true", (
        "a rename out of the app must keep the E2E marker on 'true': the pull request "
        f"removed application code. Markers: {values}"
    )
    assert values["codeql_scope_changed"] == "true"


# --- The contract between the policy, the gate and the workflow ------------


def test_the_policy_declares_one_lane_per_surface() -> None:
    """Without a lane, the skip of the job is undeclared and `required` fails."""
    lanes = {(lane["output"], lane["event"], lane["gate_job"]) for lane in _lanes()}
    required = _policy()["required_jobs"]
    for marker, job in SURFACE_JOBS.items():
        assert (marker, "pull_request", "pr-gate") in lanes, (
            f"the policy must declare a lane for {marker}: without it the skip of the "
            "job would be undeclared and `required` would fail"
        )
        lane = next(lane for lane in _lanes() if lane["output"] == marker)
        assert job in lane["skips"], f"the lane of {marker} must skip {job}"
    for lane in _lanes():
        for job in lane["skips"]:
            assert job in required, f"{job} is skipped by a lane but is not a required job"


def test_the_pr_gate_publishes_every_marker_the_policy_declares() -> None:
    """The marker has to cross the `workflow_call` boundary to reach `ci.yml`."""
    job_outputs = _pr_gate()["jobs"]["pr-gate"]["outputs"]
    call_outputs = (_pr_gate().get(True) or _pr_gate().get("on"))["workflow_call"]["outputs"]
    for lane in _lanes():
        marker = lane["output"]
        assert marker in job_outputs, f"the pr-gate job must publish {marker}"
        assert marker in call_outputs, f"the workflow_call must publish {marker}"


def test_the_markers_start_as_true() -> None:
    """Fail-closed: a diff that could not be computed leaves the heavy jobs running."""
    runs = "\n".join(step.get("run", "") for step in _pr_gate()["jobs"]["pr-gate"].get("steps", []))
    for lane in _lanes():
        assert f"{lane['output']}=true" in runs, (
            f"the pr-gate step must initialize {lane['output']} as true: the marker is "
            "the permission to skip, and the permission is not granted by default"
        )


def test_every_lane_is_read_by_the_job_it_skips() -> None:
    """A marker nobody reads is a marker that skips nothing."""
    jobs = _workflow()["jobs"]
    for lane in _lanes():
        marker = f"needs.pr-gate.outputs.{lane['output']}"
        for job in lane["skips"]:
            assert marker in jobs[job]["if"], (
                f"`{job}` must read {marker}: without it the job runs on every pull "
                "request and the lane saves nothing"
            )


def test_the_nightly_runs_the_surface_jobs_again() -> None:
    """The release workflow has never run (the repo has no tags), so the nightly is the net."""
    jobs = _workflow()["jobs"]
    assert "github.event_name == 'schedule'" in jobs["e2e-smoke"]["if"], (
        "e2e-smoke must run on the nightly: it is the only net left for the E2E surface"
    )
    assert "github.event_name == 'schedule'" in jobs["codeql"]["if"]
    assert "e2e-smoke" not in _policy()["events"]["schedule"]["accepted_skips"], (
        "the nightly executes e2e-smoke, so the policy must not accept its skip there"
    )
