"""The release ci gate reads green from the queue, not from a push run.

Every commit that lands on ``main`` arrives with a green ``ci`` job on the
exact SHA, produced by the merge queue (``event: merge_group``). The nightly
schedule (``event: schedule``) also carries a green verdict on the same SHA.
A push or pull_request run on the same SHA must not certify a tag: with the
queue in place it does not exist, and a surviving one is a stray or a
pre-queue artefact — never the verdict that authorised the merge.

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
    printf '%s' "$GH_TEST_API_RESPONSE" | jq -c "$jq_expr"
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


def test_merge_group_success_passes(tmp_path: Path) -> None:
    """A1 — green merge_group run is the canonical release verdict."""
    result = _run_gate(
        tmp_path,
        [
            {
                "name": "ci",
                "event": "merge_group",
                "conclusion": "success",
                "run_started_at": "2026-10-15T10:00:00Z",
            }
        ],
    )
    assert result.returncode == 0, result.stderr
    assert "green" in result.stdout


def test_schedule_success_passes(tmp_path: Path) -> None:
    """A2 — green schedule run (nightly mutation) also counts."""
    result = _run_gate(
        tmp_path,
        [
            {
                "name": "ci",
                "event": "schedule",
                "conclusion": "success",
                "run_started_at": "2026-10-14T03:00:00Z",
            }
        ],
    )
    assert result.returncode == 0, result.stderr
    assert "green" in result.stdout


def test_push_run_does_not_satisfy_the_gate(tmp_path: Path) -> None:
    """A3 — a green push run on the same SHA is not the release verdict.

    With the queue, push runs on ``main`` no longer exist. A surviving
    push run is either a stray or a pre-queue artefact, and must not
    certify a tag.
    """
    result = _run_gate(
        tmp_path,
        [
            {
                "name": "ci",
                "event": "push",
                "conclusion": "success",
                "run_started_at": "2026-10-15T09:00:00Z",
            }
        ],
    )
    assert result.returncode == 1, result.stderr
    assert "ci gate" in result.stderr


def test_no_runs_fails(tmp_path: Path) -> None:
    """A4 — no completed runs on the SHA fails."""
    result = _run_gate(tmp_path, [])
    assert result.returncode == 1, result.stderr
    assert "no completed" in result.stderr


def test_non_success_conclusion_fails(tmp_path: Path) -> None:
    """A5 — a merge_group run that did not pass fails the gate."""
    result = _run_gate(
        tmp_path,
        [
            {
                "name": "ci",
                "event": "merge_group",
                "conclusion": "failure",
                "run_started_at": "2026-10-15T10:00:00Z",
            }
        ],
    )
    assert result.returncode == 1, result.stderr
    assert "failure" in result.stderr


def test_api_error_fails(tmp_path: Path) -> None:
    """A6 — a GitHub API error fails the gate (no green evidence).

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
