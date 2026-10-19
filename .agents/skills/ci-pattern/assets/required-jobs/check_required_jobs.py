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
so a downstream cascade skip is never mistaken for the cause: the cause leads
the verdict (``blocked by:``) and the cascade is a single line (#787).

The policy may carry an optional ``relaxation`` field (``reason`` and
``reference``, both non-empty strings). It is validated whenever present:
a malformed field never sits in a policy waiting to cover a later
relaxation.

The policy may also carry an optional ``dependencies`` object mapping each
required job to its direct upstream jobs. When a skipped job's transitive
upstream closure intersects ``accepted_skips`` for the event, the skip is
treated as a cascade of a declared skip (cause→consequence) and the gate
accepts it (#853). Self-loops and cycles fail closed at load time, never
at verdict time.

The policy may also carry an optional ``docs_only_lane`` object naming the
event that admits the lane, the gate job that publishes the marker, the
output key that holds it and the skips it covers. The skips are accepted
only when the gate proved the marker is exactly ``"true"`` on the declared
event; any other value, a missing output, or a different event keeps every
skip undeclared and the gate fails closed (#857). The lane is purely
additive: cascade skip from ``dependencies`` (#853) and conditional skip
from ``docs_only_lane`` coexist, and neither widens ``accepted_skips``.

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
import sys
from pathlib import Path

SUCCESS = "success"
SKIPPED = "skipped"
ROOT_CAUSE_CONCLUSIONS = ("failure", "cancelled")

POLICY_KEYS = frozenset({"required_jobs", "events"})
OPTIONAL_POLICY_KEYS = frozenset({"relaxation", "dependencies", "docs_only_lane"})
EVENT_KEYS = frozenset({"accepted_skips"})
RELAXATION_KEYS = frozenset({"reason", "reference", "actor", "issue"})
RELAXATION_REQUIRED = frozenset({"reason", "reference"})
DOCS_ONLY_LANE_KEYS = frozenset({"event", "gate_job", "output", "skips"})


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


def validate_events(events: object, jobs: list, label: str) -> None:
    """Every event is an object with exactly the accepted-skips key, naming known jobs."""
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


def validate_relaxation(field: object, label: str) -> None:
    """A relaxation field carries reason, reference and a verifiable identity (HR-33)."""
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


def validate_docs_only_lane(lane: object, jobs: list, events: dict, label: str) -> None:
    """The docs-only lane names its event, the gate job, its marker and the skips it covers.

    The skips are conditional on the marker, so the lane never widens
    ``accepted_skips``; every reference must point at a declared job or event.
    Raises :class:`GateError` on every structural doubt; the ``# noqa: TRY003``
    on each raise keeps the ratchet from growing the baseline (HR-15) for what
    is the same ``{label}: {body}`` pattern the rest of the file already uses
    for every other gate error (#857).
    """
    if not isinstance(lane, dict) or set(lane) != DOCS_ONLY_LANE_KEYS:
        raise GateError(  # noqa: TRY003
            f"{label}: docs_only_lane must be an object with exactly "
            f"{sorted(DOCS_ONLY_LANE_KEYS)}")
    for key in ("event", "gate_job", "output"):
        if not isinstance(lane[key], str) or not lane[key].strip():
            raise GateError(  # noqa: TRY003
                f"{label}: docs_only_lane.{key} must be a non-empty string")
    skips = lane["skips"]
    if not isinstance(skips, list) or not skips or not all(
            isinstance(j, str) and j for j in skips):
        raise GateError(  # noqa: TRY003
            f"{label}: docs_only_lane.skips must be a non-empty list of job names")
    if len(set(skips)) != len(skips):
        raise GateError(  # noqa: TRY003
            f"{label}: docs_only_lane.skips contains duplicate job names")
    if lane["event"] not in events:
        raise GateError(  # noqa: TRY003
            f"{label}: docs_only_lane.event '{lane['event']}' is not a declared event")
    if lane["gate_job"] not in jobs or lane["gate_job"] in skips:
        raise GateError(  # noqa: TRY003
            f"{label}: docs_only_lane.gate_job '{lane['gate_job']}' must be a required job "
            "that the lane does not skip")
    unknown = set(skips) - set(jobs)
    if unknown:
        raise GateError(  # noqa: TRY003
            f"{label}: docs_only_lane.skips outside required_jobs: {sorted(unknown)}")


def validate_dependencies(field: object, jobs: list, label: str
                          ) -> dict[str, tuple[str, ...]]:
    """Optional ``dependencies``: ``{downstream: [upstream, ...]}``.

    Every name must be a non-empty string that names a required job; a
    self-loop or any cycle fails closed at load time so the verdict phase
    never sees a malformed graph (#853). The returned structure is the
    graph the verdict walks: each key is a downstream job, each value is
    its tuple of direct upstreams.
    """
    if not isinstance(field, dict):
        raise GateError(  # noqa: TRY003
            f"{label}: dependencies must be an object")
    known = set(jobs)
    normalized = _normalize_dependencies(field, known, label)
    _assert_no_cycles(normalized, label)
    return normalized


def _normalize_dependencies(field: dict, known: set[str], label: str
                            ) -> dict[str, tuple[str, ...]]:
    """Validate every entry of the ``dependencies`` field and collect its tuple.

    The per-entry checks live in :func:`_validate_dependency_entry`; this
    orchestrator only walks the field once and keeps the verdict graph
    flat so :func:`_assert_no_cycles` can iterate without indexing
    exceptions.
    """
    normalized: dict[str, tuple[str, ...]] = {}
    for job, upstreams in field.items():
        normalized[job] = _validate_dependency_entry(job, upstreams, known, label)
    return normalized


def _validate_dependency_entry(job: object, upstreams: object,
                                known: set[str], label: str) -> tuple[str, ...]:
    """One entry of the ``dependencies`` graph: key, value and each upstream.

    Raises :class:`GateError` with the specific message the verdict
    prints; the ``# noqa: TRY003`` keeps the ratchet from growing the
    baseline (HR-15) for what is the same ``{label}: {body}`` pattern the
    rest of the file already uses for every other gate error.
    """
    if not isinstance(job, str) or not job:
        raise GateError(  # noqa: TRY003
            f"{label}: dependency key must be a non-empty string")
    if job not in known:
        raise GateError(  # noqa: TRY003
            f"{label}: dependency for unknown job '{job}' (not in required_jobs)")
    if not isinstance(upstreams, list) or not all(
            isinstance(u, str) and u for u in upstreams):
        raise GateError(  # noqa: TRY003
            f"{label}: dependencies['{job}'] must be a list of job name strings")
    for upstream in upstreams:
        if upstream not in known:
            raise GateError(  # noqa: TRY003
                f"{label}: dependency '{job}' names unknown upstream "
                f"'{upstream}' (not in required_jobs)")
        if upstream == job:
            raise GateError(  # noqa: TRY003
                f"{label}: dependency '{job}' lists itself as upstream")
    return tuple(upstreams)


def _assert_no_cycles(normalized: dict[str, tuple[str, ...]], label: str) -> None:
    """Reject any cycle in the validated graph at load time (#853).

    DFS that flags any revisit on the current call stack; ``visited``
    carries the closure of nodes whose subtree has already been walked,
    ``on_stack`` is the current path. The graph is the validated one from
    :func:`_normalize_dependencies`, so a self-loop never reaches this
    step.
    """
    visited: set[str] = set()
    on_stack: set[str] = set()

    def dfs(node: str) -> None:
        if node in on_stack:
            raise GateError(  # noqa: TRY003
                f"{label}: dependencies form a cycle through '{node}'")
        if node in visited:
            return
        on_stack.add(node)
        for upstream in normalized.get(node, ()):
            dfs(upstream)
        on_stack.discard(node)
        visited.add(node)

    for job in normalized:
        dfs(job)


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
    validate_events(events, jobs, label)
    if "docs_only_lane" in data:
        # Validated whenever present, base copy or not: a malformed field must
        # not sit in the policy waiting to "cover" a later lane skip (#857).
        validate_docs_only_lane(data["docs_only_lane"], jobs, events, label)
    if "relaxation" in data:
        # Validated whenever present, base copy or not: a malformed field must
        # not sit in the policy waiting to "cover" a later relaxation.
        validate_relaxation(data["relaxation"], label)
    if "dependencies" in data:
        # Validated and normalized when present: a malformed graph never
        # sits in the policy waiting to "cover" a later cascade (#853).
        data["dependencies"] = validate_dependencies(
            data["dependencies"], jobs, label)
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


def load_needs(path: str | None) -> tuple[dict, dict]:
    """Load the serialized needs object; return job -> conclusion and job -> outputs.

    ``outputs`` carries the per-job ``outputs`` map the runner publishes
    (the lane marker lives there, #857). A bare conclusion string or an
    object without ``outputs`` is read with an empty ``outputs`` map.
    """
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
    outputs = {job: value["outputs"] if isinstance(value, dict) and isinstance(value.get("outputs"), dict) else {}
               for job, value in data.items()}
    return {job: job_conclusion(job, value) for job, value in data.items()}, outputs


def same_file(first: str, second: str) -> bool:
    """True when both paths name one file; an unanswerable question raises."""
    try:
        return Path(first).samefile(second)
    except OSError as exc:
        raise GateError(f"cannot tell the base policy from the candidate: {exc}") from exc


def load_base_policy(base_path: str | None, policy_path: str) -> dict | None:
    """The base-branch policy, or None when no comparison was requested.

    Asked for and unusable is a doubt, never a silent skip of the check.
    """
    if base_path is None:
        return None
    base = load_policy(base_path, label="base policy file")
    if same_file(base_path, policy_path):
        raise GateError(
            "--base-policy and --policy are the same file; "
            "a policy compared with itself never shows a relaxation"
        )
    return base


def relaxation_identity(policy: dict) -> tuple[str, str] | None:
    """The policy's relaxation field as a comparable value, or None if absent."""
    field = policy.get("relaxation")
    if field is None:
        return None
    return (field["reason"].strip(), field["reference"].strip())


def find_relaxations(candidate: dict, base: dict) -> list[str]:
    """List every way the candidate policy is laxer than the base copy.

    Tightening — a job added, an accepted skip removed, an event removed — is
    not listed: it needs no justification. The lane walk lives in
    :func:`_docs_only_lane_relaxations` so the body keeps its complexity under
    C901 (#857).
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
    found.extend(_docs_only_lane_relaxations(candidate, base))
    return found


def _docs_only_lane_relaxations(candidate: dict, base: dict) -> list[str]:
    """Lane-level relaxations: lane source changed or any skip added (#857).

    Split out of :func:`find_relaxations` to keep that function under C901.
    A change of the lane source (``event``, ``gate_job`` or ``output``) is a
    new relaxation even if the skips themselves are inherited.
    """
    lane = candidate.get("docs_only_lane")
    if lane is None:
        return []
    base_lane = base.get("docs_only_lane")
    before = set(base_lane["skips"]) if base_lane else set()
    found: list[str] = []
    if base_lane and any(base_lane[k] != lane[k] for k in ("event", "gate_job", "output")):
        found.append("docs-only lane source changed (event, gate_job or output)")
    for job in lane["skips"]:
        if job not in before:
            found.append(f"docs-only skip '{job}' added")
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


def policy_violations(needs: dict, policy: dict, event: str) -> tuple[list[str], set[str]]:
    """Violations of the policy by the needs keys, and the skips accepted for the event."""
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
    return violations, accepted


def docs_only_skips(policy: dict, event: str, outputs: dict) -> set[str]:
    """Skips the docs-only lane covers: only on its event, only when the gate proved docs_changed.

    The marker must be exactly ``"true"``; any other value, or a missing gate
    output, keeps every skip undeclared. The lane is purely conditional and
    never widens ``accepted_skips`` (#857).
    """
    lane = policy.get("docs_only_lane")
    if lane is None or event != lane["event"]:
        return set()
    gate_outputs = outputs.get(lane["gate_job"], {})
    if gate_outputs.get(lane["output"]) != "true":
        return set()
    return set(lane["skips"])


def classify_conclusions(needs: dict, accepted: set[str],
                          dependencies: dict[str, tuple[str, ...]] | None = None
                          ) -> dict[str, list[str]]:
    """Split the non-success conclusions into root causes, skips and unknown ones.

    A skipped job whose transitive upstream closure (over the validated
    ``dependencies`` graph) names a job in ``accepted`` is treated as a
    cascade of a declared skip and moves to ``declared_skips`` (#853). The
    cascade acceptance is purely additive: a skipped job with no upstream
    in ``accepted`` still lands in ``undeclared_skips``, and a job whose
    upstream is a ``failure`` or ``cancelled`` is the root cause (handled
    by ``policy_violations`` through ``evaluate``) and never rescues the
    cascade.
    """
    deps = dependencies or {}
    sorted_out: dict[str, list[str]] = {"root_causes": [], "undeclared_skips": [],
                                        "declared_skips": [], "unknown": []}
    for job, conclusion in sorted(needs.items()):
        if conclusion == SUCCESS:
            continue
        if conclusion == SKIPPED:
            if job in accepted or _cascade_is_accepted(job, accepted, deps):
                sorted_out["declared_skips"].append(job)
            else:
                sorted_out["undeclared_skips"].append(job)
        elif conclusion in ROOT_CAUSE_CONCLUSIONS:
            sorted_out["root_causes"].append(f"{job} ({conclusion})")
        else:
            sorted_out["unknown"].append(f"job '{job}' has unknown conclusion '{conclusion}'")
    return sorted_out


def _cascade_is_accepted(job: str, accepted: set[str],
                          dependencies: dict[str, tuple[str, ...]]) -> bool:
    """True when any transitive upstream of ``job`` is in ``accepted``.

    ``dependencies`` is the validated, cycle-free graph from the policy;
    ``seen`` bounds the walk to one visit per node.
    """
    seen: set[str] = set()
    stack = list(dependencies.get(job, ()))
    while stack:
        node = stack.pop()
        if node in seen:
            continue
        seen.add(node)
        if node in accepted:
            return True
        stack.extend(dependencies.get(node, ()))
    return False


def evaluate(needs: dict, policy: dict, event: str, base: dict | None = None,
             outputs: dict | None = None) -> dict:
    """Compare the needs object against the policy and return the verdict.

    ``base`` is the base-branch copy of the policy; None means the comparison
    was not requested, which the report records instead of assuming it clean.
    ``outputs`` carries the per-job ``outputs`` map the runner publishes; the
    docs-only lane reads its marker from there (#857). When absent, the lane
    contributes no skips.
    """
    required = policy["required_jobs"]
    violations, accepted = policy_violations(needs, policy, event)
    dependencies = policy.get("dependencies") or {}
    lane_skips = docs_only_skips(policy, event, outputs or {})
    conclusions = classify_conclusions(needs, accepted | lane_skips, dependencies)
    violations += conclusions["unknown"]
    # HR-15: un run sin ningún success entre los jobs requeridos no es un
    # PASS (todo saltado-aceptado, o ninguna señal verde que proteger).
    if required and not any(needs.get(job) == SUCCESS for job in required):
        violations.append(
            "no required job reported success: a run without a single green "
            "required job is not a pass (HR-15)")
    comparison = None if base is None else compare_with_base(policy, base)
    refused = comparison is not None and comparison["refusal"] is not None
    root_causes = conclusions["root_causes"]
    undeclared_skips = conclusions["undeclared_skips"]
    return {
        "event": event,
        "root_causes": root_causes,
        "undeclared_skips": undeclared_skips,
        "declared_skips": conclusions["declared_skips"],
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
    if report["root_causes"]:
        lines.append(f"blocked by: {', '.join(report['root_causes'])}")
    lines += render_base_comparison(report["base_comparison"])
    if report["root_causes"]:
        lines.append("root-cause failures (fix these first):")
        lines += [f"  {item}" for item in report["root_causes"]]
    if report["undeclared_skips"]:
        lines.append("skips in cascade, not declared for this event (consequences, not causes):")
        lines.append(f"  {', '.join(report['undeclared_skips'])} (skipped)")
    if report["violations"]:
        lines.append("policy violations:")
        lines += [f"  {item}" for item in report["violations"]]
    if report["declared_skips"]:
        lines.append("declared skips (accepted):")
        lines += [f"  {job} (skipped)" for job in report["declared_skips"]]
    return "\n".join(lines)


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
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
        base = load_base_policy(args.base_policy, args.policy)
        needs, outputs = load_needs(args.needs_file)
        report = evaluate(needs, policy, args.event, base, outputs)
    except GateError as exc:
        print(f"FAIL fail-closed: {exc}", file=sys.stderr)
        return 1
    print(render(report))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
