"""The release CI gate reads green from the most recent ``ci.yml`` run for the SHA.

Issue #857, PR-C1c: the gate's contract switches to the gentle-ai criterion.
The previous version filtered by ``event in {merge_group, schedule}``, which
only made sense with the merge queue. After the queue is gone, every event
that the workflow listens to (``pull_request``, ``push``, ``schedule``,
``workflow_dispatch``) is a legitimate verdict for the exact SHA — and the
*most recent* run is the only one that matters, because re-runs exist:

  1. A run that is still ``in_progress`` (or ``queued``, or any other
     non-``completed`` status) fails with a message telling the operator to
     wait and re-run the release from the same tag.
  2. A most-recent run that concluded ``failure`` fails even if an older
     run on the same SHA was green: a re-run invalidates the earlier
     verdict.
  3. A most-recent run that concluded ``success`` passes — any event,
     any name, just the file ``ci.yml`` on the exact SHA.
  4. No runs at all for the SHA fails with a message telling the operator
     to push the commit to a branch that runs ``ci.yml`` and re-run the
     release from the same tag.
  5. A GitHub API error fails the gate: no green evidence, no release.

The script under test is ``scripts/require-ci-success.sh``; the GitHub CLI is
replaced with a shim on ``PATH`` that returns a fixed API payload and applies
the script's own ``--jq`` selector with a real ``jq`` binary, so the script
exercises its real decision path against deterministic fixtures.
"""

from __future__ import annotations

import json
import os
import stat
import subprocess
from pathlib import Path
from typing import Final

REPO_ROOT: Final[Path] = Path(__file__).resolve().parents[1]
SCRIPT: Final[Path] = REPO_ROOT / "scripts" / "require-ci-success.sh"
SHA: Final[str] = "0123456789abcdef0123456789abcdef01234567"
GATE_TIMEOUT_SECONDS: Final[int] = 30


SHIM_TEMPLATE = """#!/usr/bin/env bash
# Test shim for the ``gh`` CLI. Replays a recorded API response and applies
# the script's own ``--jq`` selector with a real ``jq`` binary, so the
# script exercises its real decision path against deterministic fixtures.
#
# ``-r`` is load-bearing: ``gh api --jq`` in production formats strings
# with gojq's raw output (no JSON quoting, no escape sequences), and the
# script splits the result with ``cut -fN`` on real tab characters. Using
# ``jq -c`` instead would JSON-encode the string — tabs come out as
# literal ``\t`` and ``cut`` sees the whole row as a single field.
if [[ "$1" == "api" ]]; then
  shift
  jq_expr=""
  api_args=()
  while [[ $# -gt 0 ]]; do
    case "$1" in
      --jq) jq_expr="$2"; shift 2 ;;
      --paginate) shift ;;
      *) api_args+=("$1"); shift ;;
    esac
  done
  if [[ -n "$GH_TEST_API_RESPONSE" && -n "$jq_expr" ]]; then
    printf '%s' "$GH_TEST_API_RESPONSE" | jq -r "$jq_expr"
  fi
fi
"""


def _write_gh_shim(directory: Path) -> None:
    shim = directory / "gh"
    shim.write_text(SHIM_TEMPLATE, encoding="utf-8")
    shim.chmod(shim.stat().st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)


def _api_payload(runs: list[dict[str, object]] | None) -> str:
    body = {"total_count": len(runs or []), "workflow_runs": runs or []}
    return json.dumps(body)


def _run_gate(
    tmp_path: Path, runs: list[dict[str, object]] | None
) -> subprocess.CompletedProcess[str]:
    shim_dir = tmp_path / "shim"
    shim_dir.mkdir()
    _write_gh_shim(shim_dir)
    env = os.environ.copy()
    env["PATH"] = f"{shim_dir}{os.pathsep}{env['PATH']}"
    env["GITHUB_REPOSITORY"] = "owner/repo"
    env["GH_TEST_API_RESPONSE"] = _api_payload(runs)
    # Make sure no stray env var feeds the script a different SHA.
    env.pop("GITHUB_SHA", None)
    return subprocess.run(
        ["bash", str(SCRIPT), SHA],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
        env=env,
        cwd=str(REPO_ROOT),
        timeout=GATE_TIMEOUT_SECONDS,
    )


def test_green_passes(tmp_path: Path) -> None:
    """Rule 3: a most-recent run that concluded ``success`` passes.

    No event filter: any run of ``ci.yml`` on the exact SHA — pull_request,
    push, schedule, workflow_dispatch — is a legitimate verdict.
    """
    result = _run_gate(
        tmp_path,
        [
            {
                "created_at": "2026-10-15T10:00:00Z",
                "id": 1,
                "status": "completed",
                "conclusion": "success",
                "html_url": "https://github.com/owner/repo/actions/runs/1",
            }
        ],
    )
    assert result.returncode == 0, result.stderr
    assert "succeeded" in result.stdout


def test_in_progress_run_fails(tmp_path: Path) -> None:
    """Rule 1: a run that is still in_progress (queued, etc.) fails with
    a message that says to wait and re-run the release from the same tag."""
    result = _run_gate(
        tmp_path,
        [
            {
                "created_at": "2026-10-15T10:00:00Z",
                "id": 1,
                "status": "in_progress",
                "conclusion": None,
                "html_url": "https://github.com/owner/repo/actions/runs/1",
            }
        ],
    )
    assert result.returncode == 1, result.stderr
    assert "still in_progress" in result.stderr
    assert "Wait for it to finish" in result.stderr
    assert "re-run this release from the same tag" in result.stderr


def test_most_recent_red_with_previous_green_fails(tmp_path: Path) -> None:
    """Rule 2: the most recent run decides, even if an earlier run was green.

    A subsequent re-run that concluded ``failure`` invalidates the older
    green verdict. The gate looks at the newest run by ``created_at``
    (tie-break by ``id``) — never at the first one, never at any run that
    is not the most recent.
    """
    result = _run_gate(
        tmp_path,
        [
            {
                "created_at": "2026-10-15T10:00:00Z",
                "id": 1,
                "status": "completed",
                "conclusion": "success",
                "html_url": "https://github.com/owner/repo/actions/runs/1",
            },
            {
                "created_at": "2026-10-15T11:00:00Z",
                "id": 2,
                "status": "completed",
                "conclusion": "failure",
                "html_url": "https://github.com/owner/repo/actions/runs/2",
            },
        ],
    )
    assert result.returncode == 1, result.stderr
    assert "concluded failure" in result.stderr
    assert "Fix it on main" in result.stderr
    assert "move the tag and release again" in result.stderr


def test_no_runs_fails(tmp_path: Path) -> None:
    """Rule 4: no runs at all for the SHA fails with a message that says
    to push the commit to a branch that runs ci.yml and re-run from the tag."""
    result = _run_gate(tmp_path, [])
    assert result.returncode == 1, result.stderr
    assert "no ci run exists" in result.stderr
    assert "Push the commit" in result.stderr
    assert "re-run this release from the same tag" in result.stderr


def test_api_error_fails(tmp_path: Path) -> None:
    """Rule 5: a GitHub API error fails the gate (no green evidence, no release).

    The shim writes to stderr to mimic ``gh api`` on a network error and
    exits non-zero, which is the contract the real CLI follows.
    """
    shim_dir = tmp_path / "shim"
    shim_dir.mkdir()
    (shim_dir / "gh").write_text(
        "#!/usr/bin/env bash\necho 'gh: API error' 1>&2\nexit 1\n",
        encoding="utf-8",
    )
    (shim_dir / "gh").chmod(0o755)
    env = os.environ.copy()
    env["PATH"] = f"{shim_dir}{os.pathsep}{env['PATH']}"
    env["GITHUB_REPOSITORY"] = "owner/repo"
    env.pop("GITHUB_SHA", None)
    result = subprocess.run(
        ["bash", str(SCRIPT), SHA],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
        env=env,
        cwd=str(REPO_ROOT),
        timeout=GATE_TIMEOUT_SECONDS,
    )
    assert result.returncode == 1, result.stderr
    assert "ci gate" in result.stderr


def test_obsolete_e2e_comment_is_removed() -> None:
    """A7 — release.yml no longer claims the e2e job is non-blocking.

    The comment at ``release.yml:62-79`` described a now-removed ``|| true``
    wrapper and a no-signal "plumbing probe" rationale. The job is strict
    (``needs: preflight`` + ``needs: e2e`` on ``publish``), so the comment
    is obsolete and must be rewritten to describe what the job actually
    does.
    """
    release_yml = (REPO_ROOT / ".github" / "workflows" / "release.yml").read_text(encoding="utf-8")
    # The exact obsolete strings the comment carried, by line. None of them
    # may survive: the job is strict and there is no ``|| true`` wrapper.
    forbidden = (
        "Non-blocking Playwright smoke",
        "wraps the test runner with `|| true`",
        "Failing it would push release back without a real signal",
    )
    for needle in forbidden:
        assert needle not in release_yml, (
            f"release.yml still carries the obsolete phrase {needle!r}; "
            "rewrite the e2e job comment to match the strict job that exists."
        )
