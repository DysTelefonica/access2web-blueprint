#!/usr/bin/env python3
"""Fail closed unless every branch-protection-required CI job succeeded.

Issue #702: `main` branch protection previously required six checks
(`quality`, `pr-gate`, `pip-audit`, `gitleaks`, `trivy-config`,
`codeql`) configured by hand directly in GitHub, with nothing tying that
list back to the actual workflows. Renaming or deleting a job silently
desynchronizes it from reality.

Adapted from `ardelperal/APAP_WEB`'s `scripts/check_required_jobs.py`
(issue #766), which this repo's `docs/architecture.md` §CI gates names as
the validated reference implementation. The event set differs, so the
adaptation is not a byte-for-byte copy:

  - `ci.yml`'s own triggers are `pull_request` (with the six type
    events pinned in #857 PR-B), `push` (branches: `main`), `schedule`
    (nightly safety net) and `workflow_dispatch`. There is no
    `merge_group` trigger — the queue is gone with the gentle-ai model
    — and no tag-push trigger here (tags are `release.yml`'s domain
    and out of scope for this gate).
  - `needs:` cannot cross workflow files. `security.yml` and `codeql.yml`
    were made reusable (`on: workflow_call`) and are invoked directly as
    jobs inside `ci.yml` (`uses: ./.github/workflows/security.yml` /
    `codeql.yml`), which is what lets this aggregator `needs:` them and
    read `needs.<id>.result` exactly like any ordinary job.
  - `mutation` and `mutation-ratchet` are no longer in `ci.yml` (the
    gentle-ai model moves mutation out of CI; #857 PR-D), so they are
    out of `ALL_JOBS` by construction.
"""

from __future__ import annotations

import json
import os
import sys
from collections.abc import Mapping

#: Every job branch protection currently requires, by job id as it appears
#: in `ci.yml` (`security` and `codeql` are the reusable-workflow-call jobs,
#: not the inner jobs of `security.yml` / `codeql.yml`). Issue #882, PR-3:
#: `integration` joins the set — it runs the Postgres service and the
#: `-m integration` suite; a red integration gate holds the PR just like
#: a red quality gate does.
ALL_JOBS = frozenset(
    {"pr-gate", "quality", "integration", "security", "codeql", "docs-architecture"}
)

#: `pr-gate` only runs `if: github.event_name == 'pull_request'`
#: (issue #135, #857 PR-B). On every other trigger that fires `ci.yml`
#: — `push` (post-merge safety net), `schedule` (nightly) and
#: `workflow_dispatch` (manual) — `pr-gate` is skipped, and the heavy
#: jobs that `needs: pr-gate` accept the skip. `pull_request` is the
#: one event where the gate actually runs and the heavy jobs must see
#: a `success` from it. `merge_group` is gone: the merge queue left
#: with the gentle-ai model (#857 PR-D).
SKIPS_BY_EVENT = {
    "pull_request": frozenset(),
    "push": frozenset({"pr-gate"}),
    "schedule": frozenset({"pr-gate"}),
    "workflow_dispatch": frozenset({"pr-gate"}),
}


def _pin_output_encoding() -> None:
    """Make gate output independent from the runner locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def check_results(needs: Mapping[str, object], event_name: str) -> list[str]:
    """Return policy violations for the serialized GitHub ``needs`` object."""
    violations: list[str] = []
    missing = sorted(ALL_JOBS - needs.keys())
    if missing:
        violations.append(f"missing jobs: {', '.join(missing)}")

    allowed_skips = SKIPS_BY_EVENT.get(event_name)
    if allowed_skips is None:
        violations.append(f"unsupported event: {event_name or '<empty>'}")
        allowed_skips = frozenset()

    for job in sorted(ALL_JOBS & needs.keys()):
        payload = needs[job]
        if not isinstance(payload, Mapping):
            violations.append(f"{job}: malformed result payload")
            continue
        result = payload.get("result")
        if result == "success":
            continue
        if result == "skipped" and job in allowed_skips:
            continue
        violations.append(f"{job}: result={result!r}")
    return violations


def main() -> int:
    _pin_output_encoding()
    try:
        needs = json.loads(os.environ["CI_NEEDS_JSON"])
    except (KeyError, json.JSONDecodeError) as exc:
        print(f"FAIL required jobs: invalid CI_NEEDS_JSON ({exc})", file=sys.stderr)
        return 1
    if not isinstance(needs, dict):
        print("FAIL required jobs: CI_NEEDS_JSON must be an object", file=sys.stderr)
        return 1

    violations = check_results(needs, os.environ.get("CI_EVENT_NAME", ""))
    for violation in violations:
        print(f"FAIL required jobs: {violation}", file=sys.stderr)
    if violations:
        return 1
    print("required jobs: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
