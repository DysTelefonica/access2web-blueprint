#!/usr/bin/env python3
# ci-pattern asset — deploy review policy gate (HR-40/HR-41, DysTelefonica/team-skills#166)
"""Deploy review policy gate: resolution source declared, bootstrap
exemptions symmetric (HR-40, HR-41).

Reads the deploy review policy — a versioned JSON document owned by the
consumer's deploy workflow — and enforces:

- HR-40: the policy MUST declare the resolution source of the previous
  review under ``resolution_source`` — a ``kind`` among
  ``health-endpoint`` / ``release-artifact`` / ``deployment-record`` and a
  non-empty ``location``. A missing, empty or unknown source is a finding
  that names the cause — the gate never falls back to a heuristic
  (run lists, array indexes, default-branch HEAD).
- HR-41: every declared context (all of them are required) MUST carry a
  ``bootstrap_exempt`` flag, and the flag MUST be symmetric across
  contexts: one context exempt while another lacks declared coverage is a
  policy violation, not configuration.

Fail-closed: unreadable or malformed documents, unknown keys, contexts
without the ``required`` flag, or an empty context map exit 2 — the gate
never says OK about a policy it could not fully read (HR-3).

Exit codes: 0 clean · 1 findings · 2 invalid input or empty subject.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

POLICY_KEYS = frozenset({"policy_version", "resolution_source", "contexts"})
RESOLUTION_KEYS = frozenset({"kind", "location"})
RESOLUTION_KINDS = frozenset({"health-endpoint", "release-artifact", "deployment-record"})
CONTEXT_ENTRY_KEYS = frozenset({"required", "bootstrap_exempt"})


class GateError(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


def load_policy(path: str) -> dict:
    try:
        doc = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise GateError(f"policy {path!r} unreadable: {exc}") from exc
    if not isinstance(doc, dict):
        raise GateError("policy must be a JSON object")
    if not set(doc) <= POLICY_KEYS:
        raise GateError(f"policy keys must be within {sorted(POLICY_KEYS)}")
    if "policy_version" not in doc or doc["policy_version"] != 1:
        raise GateError('policy "policy_version" must be 1')
    if not isinstance(doc.get("contexts"), dict) or not doc["contexts"]:
        raise GateError('policy "contexts" must be a non-empty object (empty subject)')
    if "resolution_source" in doc and (
            not isinstance(doc["resolution_source"], dict)
            or set(doc["resolution_source"]) != RESOLUTION_KEYS):
        raise GateError('policy "resolution_source" must carry exactly '
                        f"{sorted(RESOLUTION_KEYS)} when present")
    return doc


def check_resolution_source(doc: dict, report: list) -> None:
    """HR-40: la ausencia de fuente declarada es hallazgo que nombra la causa
    (nunca fallback silencioso); la malformación estructural es exit 2."""
    if "resolution_source" not in doc:
        report.append("HR-40 the policy does not declare a resolution_source — "
                      "without a declared source the previous review cannot "
                      "resolve and the gate must fail loudly, never fall back "
                      "to a run-list heuristic")
        return
    source = doc["resolution_source"]
    if not isinstance(source["kind"], str) or source["kind"] not in RESOLUTION_KINDS:
        report.append(f"HR-40 resolution_source kind {source['kind']!r} is not "
                      f"supported (use one of {sorted(RESOLUTION_KINDS)}) — the "
                      "previous review must resolve from declared deployed state, "
                      "never from a run-list heuristic")
    if not isinstance(source["location"], str) or not source["location"].strip():
        report.append("HR-40 resolution_source has an empty location — without a "
                      "declared source the previous review cannot resolve, and "
                      "falling back to a heuristic is forbidden")


def check_symmetry(doc: dict, report: list) -> None:
    contexts = doc["contexts"]
    for name, entry in contexts.items():
        if not isinstance(entry, dict) or not set(entry) <= CONTEXT_ENTRY_KEYS:
            raise GateError(f"context {name!r} keys must be within "
                            f"{sorted(CONTEXT_ENTRY_KEYS)}")
        if "required" not in entry or not isinstance(entry["required"], bool):
            raise GateError(f"context {name!r}: \"required\" must be a boolean")
    flags = {name: entry.get("bootstrap_exempt") for name, entry in contexts.items()}
    if any(not isinstance(f, bool) for f in flags.values()):
        for name, flag in flags.items():
            if not isinstance(flag, bool):
                report.append(f"HR-41 context {name!r} has no declared bootstrap "
                              "exemption coverage — a required context without "
                              "declared coverage is a policy violation, never "
                              "an ad-hoc grant")
        return
    if len(set(flags.values())) > 1:
        exempt = sorted(n for n, f in flags.items() if f)
        bare = sorted(n for n, f in flags.items() if not f)
        report.append(f"HR-41 bootstrap exemptions are asymmetric: exempt "
                      f"{exempt} vs non-exempt {bare} — the exemption is a "
                      "single symmetric policy datum, never per-context "
                      "ad-hoc configuration")


def main(argv):
    ap = argparse.ArgumentParser(description="Deploy review policy gate (HR-40, HR-41).")
    ap.add_argument("--policy", required=True)
    args = ap.parse_args(argv)
    try:
        doc = load_policy(args.policy)
        report = []
        check_resolution_source(doc, report)
        check_symmetry(doc, report)
    except GateError as exc:
        print(f"DEPLOY-POLICY ERROR: {exc}", file=sys.stderr)
        return 2
    if report:
        print(f"DEPLOY-POLICY FAIL: {len(report)} finding(s)")
        for line in report:
            print(f"  - {line}")
        return 1
    print("DEPLOY-POLICY OK: resolution source declared and bootstrap "
          "exemptions symmetric across all required contexts.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
