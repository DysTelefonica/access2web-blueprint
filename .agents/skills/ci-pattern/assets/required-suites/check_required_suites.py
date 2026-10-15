#!/usr/bin/env python3
# ci-pattern asset — required-suites gate (DysTelefonica/team-skills#528)
"""Required-suites aggregator: every mandatory suite must run at least one test.

Reads ``--policy`` (logical suite names + per-event ``accepted_unran``)
and ``--evidence`` (the runner's liveness trio ``measured``/``subjects``/
``suites``). For each required suite the gate asks: did it appear in the
evidence with ``ran`` > 0, or is it explicitly accepted as unrun for the
current event? Anything else is a finding.

Liveness (HR-3): ``measured`` must be ``true`` and ``subjects`` > 0; a gate
that did not measure never says OK. Base comparison (HR-15, ``--base-
policy``) lists every relaxation of the candidate against the base copy
(a suite removed from ``required_suites``, an event added, a suite added
to ``accepted_unran`` of an existing event) and requires a ``relaxation``
field (``reason`` + ``reference`` non-empty, plus ``actor`` or ``issue``
per HR-33) the base copy does not already carry.

Exit codes: 0 clean · 1 fail-closed (missing evidence, zero-run outside
the accepted list, undeclared event, uncovered relaxation) · 2 invalid
input (unreadable file, malformed JSON, missing liveness proof,
duplicate names, unknown top-level keys).
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

POLICY_KEYS = frozenset({"required_suites", "events"})
OPTIONAL_POLICY_KEYS = frozenset({"relaxation"})
EVENT_KEYS = frozenset({"accepted_unran"})
RELAXATION_KEYS = frozenset({"reason", "reference", "actor", "issue"})
RELAXATION_REQUIRED = frozenset({"reason", "reference"})
EVIDENCE_KEYS = frozenset({"measured", "subjects", "suites"})
SUITE_KEYS = frozenset({"name", "ran", "passed", "skipped"})


class GateError(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


class GateFinding(Exception):
    """A definite verdict violation; maps to exit 1."""


def parse_json(raw: bytes, label: str) -> object:
    """Decode and parse one JSON input; every way it can be unreadable raises."""
    try:
        return json.loads(raw.decode("utf-8"))
    except UnicodeDecodeError as exc:
        raise GateError(f"{label} is not valid UTF-8: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise GateError(f"{label} is not valid JSON: {exc}") from exc
    except RecursionError as exc:
        raise GateError(f"{label} is nested too deeply to be read") from exc


def has_identity(field: dict) -> bool:
    """HR-33: a non-empty actor string or a positive issue integer."""
    actor = field.get("actor")
    issue = field.get("issue")
    return (isinstance(actor, str) and bool(actor.strip())) or (
        isinstance(issue, int) and not isinstance(issue, bool) and issue > 0)


def load_policy(path: str, label: str = "policy file") -> dict:
    """Load and validate a policy file; every structural doubt raises."""
    try:
        raw = Path(path).read_bytes()
    except OSError as exc:
        raise GateError(f"unreadable {label}: {exc}") from exc
    data = parse_json(raw, label)
    if not isinstance(data, dict):
        raise GateError(f"{label} must be an object with {sorted(POLICY_KEYS)} "
                        f"(and optionally {sorted(OPTIONAL_POLICY_KEYS)})")
    keys = set(data)
    if not keys >= POLICY_KEYS or keys - POLICY_KEYS - OPTIONAL_POLICY_KEYS:
        raise GateError(
            f"{label} must be an object with exactly {sorted(POLICY_KEYS)} "
            f"and optionally {sorted(OPTIONAL_POLICY_KEYS)}"
        )
    suites = data["required_suites"]
    if not isinstance(suites, list) or not suites \
            or not all(isinstance(s, str) and s for s in suites):
        raise GateError(f"{label}: required_suites must be a non-empty list of non-empty strings")
    if len(set(suites)) != len(suites):
        raise GateError(f"{label}: required_suites contains duplicate suite names")
    events = data["events"]
    if not isinstance(events, dict) or not events:
        raise GateError(f"{label}: events must be a non-empty object")
    known = set(suites)
    for event, spec in events.items():
        if not isinstance(spec, dict) or set(spec) != set(EVENT_KEYS):
            raise GateError(
                f"{label}: event '{event}' must be an object with exactly {sorted(EVENT_KEYS)}"
            )
        accepted = spec["accepted_unran"]
        if not isinstance(accepted, list) \
                or not all(isinstance(s, str) and s for s in accepted):
            raise GateError(f"{label}: event '{event}': accepted_unran must be a list of suite names")
        unknown = set(accepted) - known
        if unknown:
            raise GateError(
                f"{label}: event '{event}': accepted unran outside required_suites: "
                f"{sorted(unknown)}"
            )
    if "relaxation" in data:
        field = data["relaxation"]
        if (
            not isinstance(field, dict)
            or not set(field) <= set(RELAXATION_KEYS)
            or not set(RELAXATION_REQUIRED) <= set(field)
            or not all(isinstance(field[k], str) and field[k].strip()
                       for k in RELAXATION_REQUIRED)
            or not has_identity(field)
        ):
            raise GateError(
                f"{label}: relaxation must be an object with non-empty "
                f"{sorted(RELAXATION_REQUIRED)} plus a verifiable identity — a "
                'non-empty "actor" or a positive "issue" (HR-33)'
            )
    return data


def load_evidence(path: str) -> dict:
    """Load and validate the runner's measurement proof; liveness is enforced."""
    try:
        raw = Path(path).read_bytes()
    except OSError as exc:
        raise GateError(f"unreadable evidence file: {exc}") from exc
    data = parse_json(raw, "evidence file")
    if not isinstance(data, dict) or set(data) != EVIDENCE_KEYS:
        raise GateError(f"evidence must carry exactly {sorted(EVIDENCE_KEYS)}")
    if data["measured"] is not True:
        raise GateError(
            'evidence must declare "measured": true — the runner must attest a '
            "successful run; anything else is liveness exit 2"
        )
    if isinstance(data["subjects"], bool) or not isinstance(data["subjects"], int) \
            or data["subjects"] <= 0:
        raise GateError(
            'evidence "subjects" must be an integer > 0 (the count of measured '
            "subjects); zero or missing is liveness exit 2"
        )
    suites = data["suites"]
    if not isinstance(suites, list):
        raise GateError("evidence suites must be a list")
    by_name: dict[str, dict] = {}
    for i, entry in enumerate(suites):
        if not isinstance(entry, dict) or set(entry) != SUITE_KEYS:
            raise GateError(f"evidence suite {i} must carry exactly {sorted(SUITE_KEYS)}")
        name = entry["name"]
        if not isinstance(name, str) or not name:
            raise GateError(f"evidence suite {i}: name must be a non-empty string")
        if name in by_name:
            raise GateError(f"evidence contains duplicate suite name {name!r}")
        for counter in ("ran", "passed", "skipped"):
            value = entry[counter]
            if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                raise GateError(
                    f"evidence suite {name!r}: {counter} must be a non-negative integer"
                )
        by_name[name] = entry
    return {"subjects": data["subjects"], "suites": by_name}


def same_file(first: str, second: str) -> bool:
    """True when both paths name one file; an unanswerable question raises."""
    try:
        return os.path.samefile(first, second)
    except OSError as exc:
        raise GateError(f"cannot tell the base policy from the candidate: {exc}") from exc


def relaxation_identity(policy: dict) -> tuple[str, str] | None:
    """The policy's relaxation field as a comparable value, or None if absent."""
    field = policy.get("relaxation")
    if field is None:
        return None
    return (field["reason"].strip(), field["reference"].strip())


def find_relaxations(candidate: dict, base: dict) -> list[str]:
    """List every way the candidate policy is laxer than the base copy.

    Tightening — a suite added, an accepted unrun removed, an event
    removed — is not listed: it needs no justification.
    """
    found: list[str] = []
    kept = set(candidate["required_suites"])
    for suite in base["required_suites"]:
        if suite not in kept:
            found.append(f"suite '{suite}' removed from required_suites")
    for event in sorted(candidate["events"]):
        if event not in base["events"]:
            found.append(f"event '{event}' added")
            continue
        before = set(base["events"][event]["accepted_unran"])
        for suite in candidate["events"][event]["accepted_unran"]:
            if suite not in before:
                found.append(
                    f"event '{event}': accepted unran '{suite}' added"
                )
    return found


def compare_with_base(candidate: dict, base: dict) -> dict:
    """Decide whether the candidate's relaxations against the base are covered."""
    relaxations = find_relaxations(candidate, base)
    field = relaxation_identity(candidate)
    refusal = None
    if relaxations:
        if field is None:
            refusal = "the policy carries no 'relaxation' field with 'reason' and 'reference'"
        elif field == relaxation_identity(base):
            refusal = (
                "its 'relaxation' field is the one the base policy already carries; "
                "an inherited field does not cover a new relaxation"
            )
    return {"relaxations": relaxations, "field": field, "refusal": refusal}


def evaluate(policy: dict, evidence: dict, event: str, base: dict | None = None) -> dict:
    """Compare the evidence against the policy and return the verdict.

    ``accepted_unran`` is the contract for opt-in suites: a suite listed
    there for the current event is covered whether the runner emitted an
    evidence entry for it or not (the runner either ran zero tests or did
    not run the suite at all — both are accepted). Outside that list, a
    missing entry is a finding, a ``ran == 0`` entry is a finding, and a
    ``ran > 0`` entry with ``passed == 0`` is a finding: a green-less run
    is not a pass (HR-3: a run that measured nothing green never says OK).
    """
    required = policy["required_suites"]
    if event not in policy["events"]:
        # An undeclared event is fail-closed: refusing to judge is better than
        # reading a missing acceptance list as "everything accepted".
        return {
            "event": event,
            "missing": [f"event '{event}' is not declared in the policy"],
            "zero_runs": [],
            "failed_runs": [],
            "covered": [],
            "base_comparison": None,
            "passed": False,
        }
    accepted = set(policy["events"][event]["accepted_unran"])
    by_name = evidence["suites"]
    missing: list[str] = []
    zero_runs: list[str] = []
    failed_runs: list[str] = []
    covered: list[str] = []
    for suite in required:
        if suite in accepted:
            entry = by_name.get(suite)
            if entry is None:
                covered.append(f"{suite} (accepted unran for '{event}')")
            elif entry["ran"] > 0:
                covered.append(f"{suite} (ran={entry['ran']}, "
                               f"passed={entry['passed']}, "
                               f"skipped={entry['skipped']})")
            else:
                covered.append(f"{suite} (accepted unran for '{event}')")
            continue
        if suite not in by_name:
            missing.append(suite)
            continue
        entry = by_name[suite]
        if entry["ran"] > 0:
            if entry["passed"] == 0:
                failed_runs.append(suite)
                covered.append(f"{suite} (ran={entry['ran']}, passed=0, "
                               f"skipped={entry['skipped']})")
            else:
                covered.append(f"{suite} (ran={entry['ran']}, "
                               f"passed={entry['passed']}, "
                               f"skipped={entry['skipped']})")
        else:
            zero_runs.append(suite)
            covered.append(f"{suite} (ran=0)")
    comparison = None if base is None else compare_with_base(policy, base)
    refused = comparison is not None and comparison["refusal"] is not None
    return {
        "event": event,
        "missing": missing,
        "zero_runs": zero_runs,
        "failed_runs": failed_runs,
        "covered": covered,
        "base_comparison": comparison,
        "passed": not (missing or zero_runs or failed_runs or refused),
    }


def render_base_comparison(comparison: dict | None) -> list[str]:
    """State whether the base comparison ran and list what it found."""
    if comparison is None:
        return [
            "base comparison: NOT performed (no --base-policy); "
            "relaxations of the policy were not checked"
        ]
    relaxations = comparison["relaxations"]
    if not relaxations:
        return ["base comparison: performed, no relaxation detected"]
    plural = "relaxation" if len(relaxations) == 1 else "relaxations"
    lines = [f"base comparison: performed, {len(relaxations)} {plural} detected"]
    if comparison["refusal"]:
        lines.append(f"relaxations against the base policy (REFUSED: {comparison['refusal']}):")
    else:
        reason, reference = comparison["field"]
        lines.append(
            "relaxations against the base policy "
            f"(covered by relaxation: reference={reference!r}, reason={reason!r}):"
        )
    return lines + [f"  {item}" for item in relaxations]


def render(report: dict) -> str:
    """Human-readable verdict; missing suites first, zero-runs and failed-runs after."""
    lines = [f"VERDICT: {'PASS' if report['passed'] else 'FAIL'} (event: {report['event']})"]
    lines += render_base_comparison(report["base_comparison"])
    if report["missing"]:
        lines.append("required suites with no evidence entry:")
        lines += [f"  {item}" for item in report["missing"]]
    if report["zero_runs"]:
        lines.append("required suites that ran zero tests and are not accepted for this event:")
        lines += [f"  {item}" for item in report["zero_runs"]]
    if report["failed_runs"]:
        lines.append("required suites that ran tests but passed none:")
        lines += [f"  {item}" for item in report["failed_runs"]]
    if report["covered"]:
        lines.append("covered suites (ran or accepted unran):")
        lines += [f"  {item}" for item in report["covered"]]
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="fail-closed required-suites aggregator (HR-3 liveness, HR-15 shrink-only)")
    parser.add_argument("--policy", required=True, help="path to the required-suites policy JSON")
    parser.add_argument("--evidence", required=True,
                        help="path to the runner's evidence JSON (measured/subjects/suites)")
    parser.add_argument("--event", required=True, help="current event name (github.event_name)")
    parser.add_argument("--base-policy", default=None,
                        help="path to the base-branch copy of the policy; enables the relaxation check")
    args = parser.parse_args(argv)
    # A suite or event name may hold characters the stream cannot encode (a lone
    # surrogate is valid JSON); escape them instead of dying mid-verdict.
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="backslashreplace")
    try:
        policy = load_policy(args.policy)
        base = None
        if args.base_policy is not None:
            base = load_policy(args.base_policy, label="base policy file")
            if same_file(args.base_policy, args.policy):
                raise GateError("--base-policy and --policy are the same file; "
                                "a policy compared with itself never shows a relaxation")
        evidence = load_evidence(args.evidence)
        report = evaluate(policy, evidence, args.event, base)
    except GateError as exc:
        print(f"FAIL fail-closed: {exc}", file=sys.stderr)
        return 2
    print(render(report))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())