#!/usr/bin/env python3
"""Fail closed unless every branch-protection-required CI job succeeded.

Issue #702: `main` branch protection previously required six checks
(`quality`, `review-budget`, `pip-audit`, `gitleaks`, `trivy-config`,
`codeql`) configured by hand directly in GitHub, with nothing tying that
list back to the actual workflows. Renaming or deleting a job silently
desynchronizes it from reality.

Adapted from `ardelperal/APAP_WEB`'s `scripts/check_required_jobs.py`
(issue #766), which this repo's `docs/calidad-de-codigo-y-ci.md` names as
the validated reference implementation. The event set differs, so the
adaptation is not a byte-for-byte copy:

  - `ci.yml`'s own triggers are `pull_request`, `push` (branches: `main`),
    `schedule` and `workflow_dispatch`. There is no tag-push trigger here —
    tags are `release.yml`'s domain and out of scope for this gate.
  - `needs:` cannot cross workflow files. `security.yml` and `codeql.yml`
    were made reusable (`on: workflow_call`) and are invoked directly as
    jobs inside `ci.yml` (`uses: ./.github/workflows/security.yml` /
    `codeql.yml`), which is what lets this aggregator `needs:` them and
    read `needs.<id>.result` exactly like any ordinary job.
  - `mutation` is deliberately NOT one of the required jobs: it is
    schedule/workflow_dispatch-only and was never part of branch
    protection, so it stays out of `ALL_JOBS` and keeps running
    independently.
"""

from __future__ import annotations

import json
import os
import sys
from collections.abc import Mapping

#: Every job branch protection currently requires, by job id as it appears
#: in `ci.yml` (`security` and `codeql` are the reusable-workflow-call jobs,
#: not the inner jobs of `security.yml` / `codeql.yml`).
ALL_JOBS = frozenset({"review-budget", "quality", "security", "codeql"})

#: `review-budget` only runs `if: github.event_name == 'pull_request'`
#: (issue #135) — every other supported event must accept it as a skip.
#: `quality`, `security` and `codeql` carry no `if:` guard in `ci.yml`, so
#: they run — and must succeed — on every event below, `pull_request`
#: included.
SKIPS_BY_EVENT = {
    "pull_request": frozenset(),
    "push": frozenset({"review-budget"}),
    "schedule": frozenset({"review-budget"}),
    "workflow_dispatch": frozenset({"review-budget"}),
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
