# HARNESS-PROVENANCE: adapted from ardelperal/APAP_WEB tests/test_check_required_jobs.py
"""Tests for `scripts/check_required_jobs.py` (issue #702).

Every case has a success and a failure path — the whole point of this gate
is that it fails closed, so a test suite for it that only ever exercises the
happy path would be the exact hole this gate is meant to close.
"""

from __future__ import annotations

import sys
from pathlib import Path

# Make ``scripts/`` importable without sys.path games, same convention as
# tests/test_check_workflows.py and tests/test_check_test_classification.py.
SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import check_required_jobs as gate  # noqa: E402  — sys.path tweak above


def _needs(result: str = "success") -> dict[str, dict[str, str]]:
    return {job: {"result": result} for job in gate.ALL_JOBS}


def test_pull_request_requires_every_job_green() -> None:
    """`pr-gate`, `quality`, `security` and `codeql` all run on a PR."""
    assert gate.check_results(_needs(), "pull_request") == []


def test_schedule_and_dispatch_allow_pr_gate_to_skip() -> None:
    """`pr-gate` is `if: github.event_name == 'pull_request'` in ci.yml."""
    needs = _needs()
    needs["pr-gate"]["result"] = "skipped"

    assert gate.check_results(needs, "schedule") == []
    assert gate.check_results(needs, "workflow_dispatch") == []


def test_push_allows_pr_gate_to_skip() -> None:
    """Issue #857, PR-D: `push` is the post-merge safety net on `main`.

    `pr-gate` is `if: github.event_name == 'pull_request'`, so on a
    push to `main` it skips. The push event must accept the skip
    exactly like `schedule` and `workflow_dispatch` do.
    """
    needs = _needs()
    needs["pr-gate"]["result"] = "skipped"

    assert gate.check_results(needs, "push") == []


def test_push_with_quality_skipped_is_a_violation() -> None:
    """Issue #857, PR-D: a skipped `quality` on the push safety net must not pass.

    `pr-gate` is allowed to skip on `push` (it has no PR payload to
    read there), but `quality` carries no `if:` guard — the safety net
    is only as good as the jobs that actually run.
    """
    needs = _needs()
    needs["pr-gate"]["result"] = "skipped"
    needs["quality"]["result"] = "skipped"

    assert gate.check_results(needs, "push") == ["quality: result='skipped'"]


def test_merge_group_is_an_unsupported_event() -> None:
    """Issue #857, PR-D: the merge queue is gone with the gentle-ai model.

    `merge_group` was the queue's verdict event; without the queue
    there is no merge-group run to listen to. The gate must reject it
    so a stray queue-style run cannot certify a tag.
    """
    violations = gate.check_results(_needs(), "merge_group")

    assert violations == ["unsupported event: merge_group"]


def test_pr_gate_skip_is_a_violation_on_pull_request() -> None:
    """A PR without pr-gate having run is exactly the gap this closes."""
    needs = _needs()
    needs["pr-gate"]["result"] = "skipped"

    assert gate.check_results(needs, "pull_request") == ["pr-gate: result='skipped'"]


def test_security_or_codeql_skip_is_never_allowed() -> None:
    """`security` and `codeql` carry no `if:` guard — they must always run."""
    for job in ("security", "codeql", "quality"):
        needs = _needs()
        needs[job]["result"] = "skipped"
        assert gate.check_results(needs, "push") == [f"{job}: result='skipped'"]


def test_failed_job_is_a_violation_on_every_supported_event() -> None:
    needs = _needs()
    needs["quality"]["result"] = "failure"

    assert gate.check_results(needs, "pull_request") == ["quality: result='failure'"]


def test_missing_job_fails_closed() -> None:
    needs = _needs()
    needs.pop("codeql")

    assert gate.check_results(needs, "pull_request") == ["missing jobs: codeql"]


def test_unsupported_event_fails_closed() -> None:
    """A stray event ci.yml never fires must not silently pass the gate."""
    violations = gate.check_results(_needs(), "release")

    assert violations == ["unsupported event: release"]


def test_empty_event_name_fails_closed() -> None:
    violations = gate.check_results(_needs(), "")

    assert violations == ["unsupported event: <empty>"]


def test_malformed_payload_is_reported() -> None:
    needs = _needs()
    needs["security"] = "not-a-mapping"  # type: ignore[assignment]

    assert gate.check_results(needs, "pull_request") == ["security: malformed result payload"]


def test_main_reads_env_and_exits_nonzero_on_violation(monkeypatch) -> None:
    needs = _needs()
    needs["quality"]["result"] = "failure"
    monkeypatch.setenv("CI_NEEDS_JSON", __import__("json").dumps(needs))
    monkeypatch.setenv("CI_EVENT_NAME", "pull_request")

    assert gate.main() == 1


def test_main_exits_zero_on_a_clean_run(monkeypatch) -> None:
    monkeypatch.setenv("CI_NEEDS_JSON", __import__("json").dumps(_needs()))
    monkeypatch.setenv("CI_EVENT_NAME", "pull_request")

    assert gate.main() == 0


def test_main_fails_closed_on_invalid_json(monkeypatch) -> None:
    monkeypatch.setenv("CI_NEEDS_JSON", "{not json")
    monkeypatch.setenv("CI_EVENT_NAME", "pull_request")

    assert gate.main() == 1


def test_main_fails_closed_when_needs_json_missing(monkeypatch) -> None:
    monkeypatch.delenv("CI_NEEDS_JSON", raising=False)

    assert gate.main() == 1
