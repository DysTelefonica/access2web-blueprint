#!/usr/bin/env python3
# ci-pattern asset — fail-closed required-jobs aggregator (DysTelefonica/team-skills#132, #161)
"""Aggregator gate over a serialized GitHub Actions ``needs`` object.

Reads ``toJSON(needs)`` — one key per dependency job, value = the object
``{"result": <conclusion>, "outputs": {...}}`` the runner emits; a bare
conclusion string is accepted too — plus a policy file that owns every job
name, event name and accepted skip. The script hardcodes none of them:
replacing the policy retargets the gate, so the same asset serves any consumer.

Fail-closed contract: any doubt exits non-zero. That includes unreadable or
malformed input, an undeclared event, a policy job absent from ``needs``, a
``needs`` key the policy does not cover (the workflow grew, the policy did
not), an undeclared skip, and any conclusion that is not ``success`` or a
declared skip.

Output separates root-cause failures (``failure``/``cancelled``) from skips,
so a downstream cascade skip is never mistaken for the cause.

The policy may carry an optional ``relaxation`` field (``reason`` and
``reference``, both non-empty strings). It is validated whenever present:
a malformed field never sits in a policy waiting to cover a later
relaxation.

Base comparison (HR-15): ``--base-policy`` takes the copy of the policy on the
base branch. The run is still judged with the candidate policy — adding a job
means adding it to the workflow, to ``needs`` and to the policy in one PR, so
the base copy could never judge that run — but every relaxation of the
candidate against the base (a job dropped from ``required_jobs``, an accepted
skip added, an event added) must be covered by a ``relaxation`` field the base
does not already carry. Without ``--base-policy`` the verdict says, in its
output, that the comparison was not performed.

Exit codes:
    0  every job succeeded or was a declared skip for the event, and no
       relaxation against the base policy went uncovered
    1  fail-closed: root-cause failure, policy violation, uncovered
       relaxation, or doubt
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

SUCCESS = "success"
SKIPPED = "skipped"
ROOT_CAUSE_CONCLUSIONS = ("failure", "cancelled")

POLICY_KEYS = frozenset({"required_jobs", "events"})
OPTIONAL_POLICY_KEYS = frozenset({"relaxation"})
EVENT_KEYS = frozenset({"accepted_skips"})
RELAXATION_KEYS = frozenset({"reason", "reference", "actor", "issue"})
RELAXATION_REQUIRED = frozenset({"reason", "reference"})


class GateError(Exception):
    """A condition the gate refuses to interpret; maps to exit 1."""


def parse_json(raw: bytes, label: str):
    """Decode and parse one JSON input; every way it can be unreadable raises.

    ``label`` names the input in the message. Bytes are decoded here, strictly,
    so the verdict does not depend on the locale of the runner.
    """
    try:
        return json.loads(raw.decode("utf-8"))
    except UnicodeDecodeError as exc:
        raise GateError(f"{label} is not valid UTF-8: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise GateError(f"{label} is not valid JSON: {exc}") from exc
    except RecursionError as exc:
        raise GateError(f"{label} is nested too deeply to be read") from exc


def load_policy(path: str, label: str = "policy file") -> dict:
    """Load and validate a policy file; every structural doubt raises.

    ``label`` names the file in the message, so a rejected base copy is not
    mistaken for a rejected candidate.
    """
    try:
        raw = Path(path).read_bytes()
    except OSError as exc:
        raise GateError(f"unreadable {label}: {exc}") from exc
    data = parse_json(raw, label)
    keys = set(data) if isinstance(data, dict) else set()
    if not keys >= POLICY_KEYS or keys - POLICY_KEYS - OPTIONAL_POLICY_KEYS:
        raise GateError(
            f"{label} must be an object with exactly {sorted(POLICY_KEYS)} "
            f"and optionally {sorted(OPTIONAL_POLICY_KEYS)}"
        )
    jobs = data["required_jobs"]
    if not isinstance(jobs, list) or not jobs or not all(isinstance(j, str) and j for j in jobs):
        raise GateError(f"{label}: required_jobs must be a non-empty list of non-empty strings")
    if len(set(jobs)) != len(jobs):
        raise GateError(f"{label}: required_jobs contains duplicate job names")
    events = data["events"]
    if not isinstance(events, dict):
        raise GateError(f"{label}: events must be an object")
    known = set(jobs)
    for event, spec in events.items():
        # accepted_skips is mandatory: an event that omits it would be read as
        # "no skips accepted" by one reader and as a crash by another.
        if not isinstance(spec, dict) or set(spec) != set(EVENT_KEYS):
            raise GateError(
                f"{label}: event '{event}' must be an object with exactly {sorted(EVENT_KEYS)}"
            )
        skips = spec["accepted_skips"]
        if not isinstance(skips, list) or not all(isinstance(s, str) and s for s in skips):
            raise GateError(f"{label}: event '{event}': accepted_skips must be a list of job names")
        unknown = set(skips) - known
        if unknown:
            raise GateError(
                f"{label}: event '{event}': accepted skips outside required_jobs: {sorted(unknown)}"
            )
    if "relaxation" in data:
        # Validated whenever present, base copy or not: a malformed field must
        # not sit in the policy waiting to "cover" a later relaxation.
        field = data["relaxation"]
        actor = field.get("actor") if isinstance(field, dict) else None
        issue = field.get("issue") if isinstance(field, dict) else None
        has_identity = (isinstance(actor, str) and bool(actor.strip())) or (
            isinstance(issue, int) and not isinstance(issue, bool) and issue > 0
        )
        if (
            not isinstance(field, dict)
            or not set(field) <= set(RELAXATION_KEYS)
            or not set(RELAXATION_REQUIRED) <= set(field)
            or not all(isinstance(field[k], str) and field[k].strip()
                       for k in RELAXATION_REQUIRED)
            or not has_identity
        ):
            raise GateError(
                f"{label}: relaxation must be an object with non-empty "
                f"{sorted(RELAXATION_REQUIRED)} plus a verifiable identity — a "
                'non-empty "actor" or a positive "issue" (HR-33)'
            )
    return data


def job_conclusion(job: str, value: object) -> str:
    """Read one job's conclusion from either accepted shape of its value."""
    if isinstance(value, str):
        return value
    if isinstance(value, dict) and isinstance(value.get("result"), str):
        return value["result"]
    raise GateError(
        "needs input must map each job to a conclusion string or to an object "
        f"with a string 'result' (the toJSON(needs) shape); job '{job}' has neither"
    )


def load_needs(path: str | None) -> dict:
    """Load the serialized needs object and return it as job -> conclusion."""
    try:
        if path is None:
            if sys.stdin.isatty():
                raise GateError("no input: pipe the needs JSON or pass --needs-file")
            raw = sys.stdin.buffer.read()
        else:
            raw = Path(path).read_bytes()
    except OSError as exc:
        raise GateError(f"unreadable needs input: {exc}") from exc
    if not raw.strip():
        raise GateError("needs input is empty")
    data = parse_json(raw, "needs input")
    if not isinstance(data, dict) or not data:
        raise GateError("needs input must be a non-empty object of job -> conclusion")
    if "" in data:
        raise GateError("needs input must map non-empty job names to conclusions")
    return {job: job_conclusion(job, value) for job, value in data.items()}


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

    Tightening — a job added, an accepted skip removed, an event removed — is
    not listed: it needs no justification.
    """
    found: list[str] = []
    kept = set(candidate["required_jobs"])
    for job in base["required_jobs"]:
        if job not in kept:
            found.append(f"job '{job}' removed from required_jobs")
    for event in sorted(candidate["events"]):
        if event not in base["events"]:
            # An event the base refused outright is now judged at all.
            found.append(f"event '{event}' added")
            continue
        before = set(base["events"][event]["accepted_skips"])
        for job in candidate["events"][event]["accepted_skips"]:
            if job not in before:
                found.append(f"event '{event}': accepted skip '{job}' added")
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
            # The base branch already holds this field: it justified an earlier
            # change and cannot be reused, untouched, for a new one.
            refusal = (
                "its 'relaxation' field is the one the base policy already carries; "
                "an inherited field does not cover a new relaxation"
            )
    return {"relaxations": relaxations, "field": field, "refusal": refusal}


def evaluate(needs: dict, policy: dict, event: str, base: dict | None = None) -> dict:
    """Compare the needs object against the policy and return the verdict.

    ``base`` is the base-branch copy of the policy; None means the comparison
    was not requested, which the report records instead of assuming it clean.
    """
    required = policy["required_jobs"]
    violations: list[str] = []
    if event not in policy["events"]:
        violations.append(f"event '{event}' is not declared in the policy")
        accepted: set[str] = set()
    else:
        accepted = set(policy["events"][event]["accepted_skips"])
    # HR-15: accepted_skips que cubre TODOS los jobs requeridos del evento
    # convierte al gate en algo que no puede bloquear nada; eso nunca es
    # un PASS, con o sin drift en el run concreto.
    if accepted and set(required) <= accepted:
        violations.append(
            f"accepted_skips covers every required job for event '{event}': "
            "a gate that cannot block anything is not a pass (HR-15)")

    # S3 rule: a needs key the policy does not cover means the workflow grew
    # without updating the policy; fail-closed instead of silently ignoring it.
    for job in sorted(set(needs) - set(required)):
        violations.append(f"needs key '{job}' is not covered by the policy")
    for job in required:
        if job not in needs:
            violations.append(f"required job '{job}' is absent from needs")

    root_causes: list[str] = []
    undeclared_skips: list[str] = []
    declared_skips: list[str] = []
    for job, conclusion in sorted(needs.items()):
        if conclusion == SUCCESS:
            continue
        if conclusion == SKIPPED:
            if job in accepted:
                declared_skips.append(job)
            else:
                undeclared_skips.append(job)
        elif conclusion in ROOT_CAUSE_CONCLUSIONS:
            root_causes.append(f"{job} ({conclusion})")
        else:
            violations.append(f"job '{job}' has unknown conclusion '{conclusion}'")
    # HR-15: un run sin ningún success entre los jobs requeridos no es un
    # PASS (todo saltado-aceptado, o ninguna señal verde que proteger).
    if required and not any(needs.get(job) == SUCCESS for job in required):
        violations.append(
            "no required job reported success: a run without a single green "
            "required job is not a pass (HR-15)")
    comparison = None if base is None else compare_with_base(policy, base)
    refused = comparison is not None and comparison["refusal"] is not None
    return {
        "event": event,
        "root_causes": root_causes,
        "undeclared_skips": undeclared_skips,
        "declared_skips": declared_skips,
        "violations": violations,
        "base_comparison": comparison,
        "passed": not (root_causes or undeclared_skips or violations or refused),
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
    """Human-readable verdict; root causes first, cascade skips after."""
    lines = [f"VERDICT: {'PASS' if report['passed'] else 'FAIL'} (event: {report['event']})"]
    lines += render_base_comparison(report["base_comparison"])
    if report["root_causes"]:
        lines.append("root-cause failures (fix these first):")
        lines += [f"  {item}" for item in report["root_causes"]]
    if report["undeclared_skips"]:
        lines.append("skips in cascade, not declared for this event (consequences, not causes):")
        lines += [f"  {job} (skipped)" for job in report["undeclared_skips"]]
    if report["violations"]:
        lines.append("policy violations:")
        lines += [f"  {item}" for item in report["violations"]]
    if report["declared_skips"]:
        lines.append("declared skips (accepted):")
        lines += [f"  {job} (skipped)" for job in report["declared_skips"]]
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="fail-closed required-jobs aggregator")
    parser.add_argument("--policy", required=True, help="path to the required-jobs policy JSON")
    parser.add_argument("--event", required=True, help="current event name (github.event_name)")
    parser.add_argument(
        "--needs-file",
        default=None,
        help="path to the serialized needs JSON (default: read stdin)",
    )
    parser.add_argument(
        "--base-policy",
        default=None,
        help="path to the base-branch copy of the policy; enables the relaxation check",
    )
    args = parser.parse_args(argv)
    # A job or event name may hold characters the stream cannot encode (a lone
    # surrogate is valid JSON); escape them instead of dying mid-verdict.
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="backslashreplace")
    try:
        policy = load_policy(args.policy)
        base = None
        if args.base_policy is not None:
            # Asked for and unusable is a doubt, never a silent skip of the check.
            base = load_policy(args.base_policy, label="base policy file")
            if same_file(args.base_policy, args.policy):
                raise GateError(
                    "--base-policy and --policy are the same file; "
                    "a policy compared with itself never shows a relaxation"
                )
        needs = load_needs(args.needs_file)
        report = evaluate(needs, policy, args.event, base)
    except GateError as exc:
        print(f"FAIL fail-closed: {exc}", file=sys.stderr)
        return 1
    print(render(report))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
