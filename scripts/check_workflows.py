#!/usr/bin/env python3
"""CI gate on .github/workflows/*.yml — issues #135, #141, #152.

Ten checks, applied by hand across PRs #126 / #129 / #142 and now pinned
by the gate in `tests/test_check_workflows.py`:

 1. No duplicate YAML keys (GitHub rejects them with startup_failure that
    silently disappears from the rollup).
 2. Every job declares timeout-minutes (GitHub default is 360 minutes — the
    exact failure mode that hit security-deep until PR #131 fixed it).
 3. No service container fixes a host port (ports: - 5432:5432 silently
    shares a DB between runners; vacuous today, future-proofs the next
    person who adds a service).
 4. Every uses: pinned by a 40-hex SHA (blinds against future tag drift).
 5. Every job invoking `docker run` checks the daemon first, wrapped in
    `timeout` (a bare `docker info` hangs in the very scenario it exists to
    detect). ci.yml::security had already slipped through without it.
 6. Every job declares a `concurrency` group. Missing group -> ERROR; a
    non-cancel-safe `cancel-in-progress: true` -> WARN; cancel-in-progress:
    false (or absent) is the silent default this repo chose.
 7. Workflow-level `permissions.contents` is `read` (or a list that contains
    `read`). Missing `permissions:` block -> WARN — the GITHUB_TOKEN inherits
    a restrictive default, but this is not pinned to `contents: read`.
 8. Every workflow that declares `env.PYTHON_VERSION` agrees on the value.
    The interpreter is part of the gate (Hard Rule #15 of the
    `oracle-vps-github-runners` skill); two declarations with two values
    means one half of CI ships on a Python the test half never saw.
 9. No two jobs in the same workflow share a `concurrency.group` literal.
    `gate-quality` on two jobs means the second job queues behind the
    first — FIFO at a granularity nobody asked for.
10. (carried from #135) Duplicate-key, timeout-minutes, host-port-fix,
    uses-not-pinned, docker-preflight, concurrency. The four added in #152
    are #6 (WARN branch), #7, #8, #9.

Hard Rule 1 of the deterministic-quality-harness skill: no `|| true`, no
`continue-on-error`. Exit 0 = clean (warnings allowed), exit 1 = any
violation. A WARN is not a violation; it is a printed line that the next
operator gets to act on.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import yaml

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------
WORKFLOWS_DIR = Path(__file__).resolve().parent.parent / ".github" / "workflows"

#: Recommended timeout per gate criticality. Fail on missing altogether.
_TIMEOUT_RECOMMENDED = {
    "review-budget": 5,  # PR-size + branch-name, both <1s
    "security": 15,  # docker pull + scan; observed 17-min zombie in #117
    "quality": 30,  # ruff + mypy + pytest + 4 quality gates
    "mutation": 60,  # cosmic-ray is slow by design
}

#: GitHub Actions SHA — 40 lowercase hex characters.
_SHA_PIN = re.compile(r"^[0-9a-f]{40}$")

#: An actual `timeout` INVOCATION, not the word. The preflight messages in this
#: repo say "15-minute silent timeout", so a substring test for `timeout` passes
#: on the prose and reports an unwrapped guard as wrapped — the same class of
#: false negative these checks exist to prevent.
_TIMEOUT_CALL = re.compile(r"(?:^|[|&;(`$]|\s)timeout\s+\d")

#: `ports:` value that fixes the host side, e.g. `5432:5432`. A bare
#: number or `number:` (empty host) does not.
_HOST_PORT_FIX = re.compile(r"^(?:\d{1,3}(?:\.\d{1,3}){3}|\d+):\d+$")

#: Repo-approved allowlist of jobs that may safely use `cancel-in-progress: true`.
#:
#: Empty by design (audit 2026-08-12, umbrella #117). The runner pool is two VPS;
#: every gate that consumed one already paid the cost. Adding a job here is a
#: deliberate decision that the in-flight work is cheap to discard — not a
#: stylistic preference. Tests patch this set via `monkeypatch.setattr` to
#: exercise the silent branch without modifying the source.
_CANCEL_SAFE_JOBS: set[str] = set()


# --------------------------------------------------------------------------------------------
# DUPLICATE-KEY DETECTION (check 1)
# --------------------------------------------------------------------------------------------
# PyYAML's safe_load silently drops duplicates (keeps the last, vanishes
# the rest from the in-memory dict). Walking the Node tree preserves every
# key/value pair with line numbers from start_mark, so we can detect what
# GitHub would reject.


def _find_duplicate_keys(text: str) -> list[tuple[int, str]]:
    """Return `[(line_no, key), ...]` for every duplicate key.

    Implementation note: PyYAML's default Composer raises on duplicates AND
    drops them silently in `safe_load`. To detect duplicates portably across
    the C and pure-Python PyYAML builds, we use a line-based scanner that
    tracks the SCOPE of each key (parent path + list-item identity).

    Two keys are duplicates iff they have the same name AND the same scope:
    same parent chain AND the same list-item (when in a list). Different
    list items (e.g. `steps:` containing multiple `- name: ...` steps) are
    distinct scopes even though their parent path matches.

    That is what GitHub rejects with `startup_failure`.
    """
    lines = text.splitlines()
    keys: list[tuple[int, int, str, bool, bool, int]] = []
    in_block_scalar = False
    block_scalar_min_indent = -1
    expecting_continuation = False
    continuation_min_indent = -1

    def _next_meaningful_indent(start: int) -> int | None:
        for j in range(start, len(lines)):
            stripped = lines[j].lstrip()
            if not stripped or stripped.startswith("#"):
                continue
            return len(lines[j]) - len(stripped)
        return None

    for i, raw in enumerate(lines):
        stripped = raw.lstrip()
        if not stripped or stripped.startswith("#"):
            in_block_scalar = False
            expecting_continuation = False
            continue
        line_indent = len(raw) - len(stripped)

        if in_block_scalar:
            if line_indent >= block_scalar_min_indent:
                continue
            in_block_scalar = False

        if expecting_continuation and line_indent > continuation_min_indent:
            continue
        expecting_continuation = False

        is_list_item = stripped.startswith("- ") or stripped == "-"
        if is_list_item:
            content = stripped[2:] if stripped.startswith("- ") else ""
            dash_indent = line_indent
        else:
            content = stripped
            dash_indent = -1

        colon_idx = -1
        for j, ch in enumerate(content):
            if ch == ":":
                if j + 1 < len(content) and content[j + 1] == " ":
                    colon_idx = j
                    break
                elif j + 1 == len(content):
                    colon_idx = j
                    break
        if colon_idx <= 0:
            continue

        key_name = content[:colon_idx].strip()
        if len(key_name) >= 2 and (
            (key_name[0] == '"' and key_name[-1] == '"')
            or (key_name[0] == "'" and key_name[-1] == "'")
        ):
            key_name = key_name[1:-1]
        if not key_name:
            continue

        target_indent = (dash_indent + 2) if is_list_item else line_indent

        value_part = content[colon_idx + 1 :].lstrip()
        starts_block = value_part.startswith("|") or value_part.startswith(">")
        if starts_block:
            in_block_scalar = True
            block_scalar_min_indent = target_indent + 2
            starts_mapping = True
        else:
            # Plain value. Scalar vs mapping is decided by what comes next.
            nxt = _next_meaningful_indent(i + 1)
            starts_mapping = nxt is not None and nxt > target_indent
            if not starts_mapping:
                # Scalar on one line — a more-indented following line is part
                # of the same scalar, not a new key.
                expecting_continuation = True
                continuation_min_indent = target_indent

        keys.append((i + 1, target_indent, key_name, starts_mapping, is_list_item, dash_indent))

    # ---- Pass 2: walk keys, build scope stack, find duplicates ----
    duplicates: list[tuple[int, str]] = []
    parent_stack: list[tuple[int, str]] = []
    current_list_item_id: int | None = None
    current_list_item_dash_indent: int | None = None
    seen_per_scope: dict[tuple[tuple[str, ...], int | None], dict[str, int]] = {((), None): {}}
    for line_no, target_indent, key_name, starts_mapping, is_list_item, dash_indent in keys:
        if is_list_item:
            current_list_item_id = line_no
            current_list_item_dash_indent = dash_indent
        else:
            if (
                current_list_item_dash_indent is not None
                and target_indent <= current_list_item_dash_indent
            ):
                current_list_item_id = None
                current_list_item_dash_indent = None

        while parent_stack and parent_stack[0][0] > target_indent:
            parent_stack.pop(0)
        if parent_stack and parent_stack[0][0] == target_indent:
            parent_stack.pop(0)

        current_path = tuple(k for _, k in parent_stack)
        scope_key = (current_path, current_list_item_id)
        seen = seen_per_scope.setdefault(scope_key, {})

        if key_name in seen:
            duplicates.append((line_no, key_name))
        else:
            seen[key_name] = line_no

        if starts_mapping:
            parent_stack.insert(0, (target_indent, key_name))

    return duplicates


# --------------------------------------------------------------------------------------------
# CHECKS 2..7 (need a parsed dict)
# --------------------------------------------------------------------------------------------
# Convention: each check yields `Finding = (severity, location, message)`
# where severity is "error" (violation -> exit 1) or "warn" (printed, exit 0).
Finding = tuple[str, str, str]


def _check_timeout_minutes(doc: dict[str, Any]) -> Iterator[Finding]:
    """Yield a violation for every job missing `timeout-minutes`."""
    jobs = doc.get("jobs") or {}
    if not isinstance(jobs, dict):
        return
    for job_name, job_def in jobs.items():
        if not isinstance(job_def, dict):
            continue
        if "timeout-minutes" in job_def:
            continue
        recommended = _TIMEOUT_RECOMMENDED.get(job_name)
        hint = (
            f" (recommend {recommended} for `{job_name}` criticality)"
            if recommended is not None
            else ""
        )
        yield (
            "error",
            f"jobs.{job_name}",
            f"job `{job_name}` does not declare `timeout-minutes`{hint}",
        )


def _check_service_ports(doc: dict[str, Any]) -> Iterator[Finding]:
    """Yield a violation for every fixed host port in `services:`."""
    jobs = doc.get("jobs") or {}
    if not isinstance(jobs, dict):
        return
    for job_name, job_def in jobs.items():
        if not isinstance(job_def, dict):
            continue
        services = job_def.get("services")
        if not isinstance(services, dict):
            continue
        for service_name, service_def in services.items():
            if not isinstance(service_def, dict):
                continue
            ports = service_def.get("ports")
            if not isinstance(ports, list):
                continue
            for entry in ports:
                if isinstance(entry, str) and _HOST_PORT_FIX.match(entry):
                    yield (
                        "error",
                        f"jobs.{job_name}.services.{service_name}.ports",
                        f"service `{service_name}` in job `{job_name}` fixes host port `{entry}`",
                    )


def _executable(script: str) -> list[str]:
    """`script` split into lines, comment-only lines dropped.

    These workflows document the commands they stopped using, so a scan that
    reads comments reports the explanation as the offence — which teaches people
    to delete explanations.
    """
    return [line for line in script.splitlines() if not line.strip().startswith("#")]


def _check_docker_preflight(doc: dict[str, Any]) -> Iterator[Finding]:
    """Yield a violation for every job reaching `docker run` unguarded.

    A wedged daemon makes `docker run` block rather than fail, so the job goes
    silent until its timeout. The guard must be wrapped in `timeout`: a bare
    `docker info` hangs in exactly the scenario it exists to detect, which makes
    it indistinguishable from no guard at all.

    The guard may sit in an earlier step OR earlier in the same step — this repo
    uses the latter — so both are accepted, and only text before the first
    `docker run` counts.
    """
    jobs = doc.get("jobs") or {}
    if not isinstance(jobs, dict):
        return
    for job_name, job_def in jobs.items():
        if not isinstance(job_def, dict):
            continue
        steps = job_def.get("steps")
        if not isinstance(steps, list):
            continue
        before: list[str] = []
        for step in steps:
            lines = (
                _executable(str((step or {}).get("run") or "")) if isinstance(step, dict) else []
            )
            index = next((i for i, line in enumerate(lines) if "docker run" in line), None)
            if index is None:
                before.extend(lines)
                continue
            guard = "\n".join([*before, *lines[:index]])
            if "docker info" not in guard:
                yield (
                    "error",
                    f"jobs.{job_name}",
                    f"job `{job_name}` reaches `docker run` with no `docker info` check",
                )
            elif not _TIMEOUT_CALL.search(guard):
                yield (
                    "error",
                    f"jobs.{job_name}",
                    f"job `{job_name}` checks `docker info` but not under `timeout`; "
                    f"an unwrapped check hangs exactly when the daemon is wedged",
                )
            break


def _check_concurrency(doc: dict[str, Any]) -> Iterator[Finding]:
    """Yield a violation for jobs without a FIFO concurrency group.

    Audit 2026-08-12 (umbrella #117) re-classified the cancel-in-progress
    policy. The literal "must be false" reading of the audit would have
    broken main, because today every job uses `false` — but the intent
    of the runner-skill rule is to surface any design where cancelling
    is allowed without an explicit repo decision. The re-classified rule:

      - missing `concurrency:` block     -> ERROR   (the runner pool is shared
                                                    and no group means no FIFO;
                                                    this is the case Hard Rule
                                                    #1 is asking to block)
      - `cancel-in-progress: true` AND
        job NOT in `_CANCEL_SAFE_JOBS`   -> WARN    (deliberate design choice,
                                                    but cancelling still
                                                    discards work that already
                                                    paid for a runner slot;
                                                    add the job name to the
                                                    allowlist to silence)
      - `cancel-in-progress: true` AND
        job in `_CANCEL_SAFE_JOBS`       -> silent  (the repo declares the
                                                    in-flight work cheap to
                                                    discard)
      - `cancel-in-progress: false` or
        absent                           -> silent  (the per-gate FIFO shape
                                                    this repo chose)

    Tests exercise the allowlist via `monkeypatch.setattr(check_workflows,
    "_CANCEL_SAFE_JOBS", {job_name})`; the production constant is empty.
    """
    jobs = doc.get("jobs") or {}
    if not isinstance(jobs, dict):
        return
    top_level = doc.get("concurrency")
    for job_name, job_def in jobs.items():
        if not isinstance(job_def, dict):
            continue
        concurrency = job_def.get("concurrency") or top_level
        if not isinstance(concurrency, dict) or not concurrency.get("group"):
            yield (
                "error",
                f"jobs.{job_name}",
                f"job `{job_name}` declares no concurrency group",
            )
            continue
        if concurrency.get("cancel-in-progress") is True and job_name not in _CANCEL_SAFE_JOBS:
            yield (
                "warn",
                f"jobs.{job_name}",
                f"job `{job_name}` uses `cancel-in-progress: true`; "
                f"cancelling discards work that already consumed one of two "
                f"runners. Add `{job_name}` to _CANCEL_SAFE_JOBS in "
                f"scripts/check_workflows.py to silence this warning.",
            )


def _check_permissions(doc: dict[str, Any]) -> Iterator[Finding]:
    """Yield a finding for the workflow-level `permissions` block.

    Audit 2026-08-12 (oracle-vps-github-runners v2.0, Hard Rule #6): the
    GITHUB_TOKEN must be pinned to `contents: read` (or a list that contains
    `read`). The default is restrictive but not pinned, so a missing block is
    a WARN — the workflow happens to inherit a restrictive default, but the
    next operator to add a job that requests `contents: write` will not get a
    gate catching the privilege escalation.

      - no `permissions:` block           -> WARN  (default is restrictive but
                                                    not pinned to `read`)
      - `permissions.contents: read`      -> silent (the minimum this gate
                                                    enforces)
      - `permissions.contents: [read, …]` -> silent (superset accepted; the
                                                    runner-skill rule is
                                                    "no less than read")
      - `permissions.contents: write`     -> ERROR
      - shorthand `permissions: read-all`
        or `permissions: write-all`       -> ERROR (grants more than the
                                                    gate's minimum)
      - non-mapping permissions           -> ERROR (unparseable for this gate)
    """
    permissions = doc.get("permissions")
    if permissions is None:
        yield (
            "warn",
            "permissions",
            "workflow declares no `permissions:` block; the GITHUB_TOKEN "
            "inherits a restrictive default, but this is not pinned to "
            "`contents: read` — add a top-level `permissions: { contents: read }`",
        )
        return
    if not isinstance(permissions, dict):
        yield (
            "error",
            "permissions",
            f"workflow uses shorthand `permissions: {permissions}`; "
            f"replace with `permissions: {{ contents: read }}` so the "
            f"GITHUB_TOKEN is pinned to the minimum-readable scope",
        )
        return
    contents = permissions.get("contents")
    if contents == "read":
        return
    if isinstance(contents, list) and "read" in contents:
        return
    yield (
        "error",
        "permissions.contents",
        f"workflow `permissions.contents` is `{contents!r}`; must be "
        f"`read` or a list that contains `read`",
    )


def _check_concurrency_group_uniqueness(doc: dict[str, Any]) -> Iterator[Finding]:
    """Yield a violation when two jobs in the same workflow share a group.

    Audit 2026-08-12 (oracle-vps-github-runners v2.0, Hard Rule #9): two
    jobs in the same workflow sharing a `concurrency.group` means a new
    push queues behind the in-flight run of the OTHER job — so the second
    job does not actually start until the first finishes. That is FIFO at
    a granularity nobody asked for, and it silently serialises the gate.

    The check is per-workflow: the same group literal in two different
    workflows is fine, because the runners and the FIFO queues are
    different. (security.yml::gitleaks and ci.yml::quality are not on the
    same machine, and even if they were, the cross-workflow collision
    would not serialise either one.)
    """
    jobs = doc.get("jobs") or {}
    if not isinstance(jobs, dict):
        return
    top_level = doc.get("concurrency")
    groups: dict[str, list[str]] = {}
    for job_name, job_def in jobs.items():
        if not isinstance(job_def, dict):
            continue
        concurrency = job_def.get("concurrency") or top_level
        if not isinstance(concurrency, dict):
            continue
        group = concurrency.get("group")
        if not isinstance(group, str) or not group:
            continue
        groups.setdefault(group, []).append(job_name)
    for group, names in groups.items():
        if len(names) > 1:
            yield (
                "error",
                "concurrency.group",
                f"jobs {names} share `concurrency.group: {group!r}`; "
                f"each job needs a distinct group so they do not queue "
                f"behind each other on the same workflow run",
            )


def _check_uses_pinned(doc: dict[str, Any]) -> Iterator[Finding]:
    """Yield a violation for every `uses:` not pinned by 40-hex SHA."""
    jobs = doc.get("jobs") or {}
    if not isinstance(jobs, dict):
        return
    for job_name, job_def in jobs.items():
        if not isinstance(job_def, dict):
            continue
        steps = job_def.get("steps")
        if not isinstance(steps, list):
            continue
        for step in steps:
            if not isinstance(step, dict):
                continue
            uses = step.get("uses")
            if not isinstance(uses, str):
                continue
            if "@" not in uses:
                yield (
                    "error",
                    f"jobs.{job_name}.steps",
                    f"step `{step.get('name', '?')}` in job `{job_name}` "
                    f"uses `{uses}` without `@<ref>`",
                )
                continue
            ref = uses.rsplit("@", 1)[1]
            if not _SHA_PIN.match(ref):
                yield (
                    "error",
                    f"jobs.{job_name}.steps",
                    f"step `{step.get('name', '?')}` in job `{job_name}` "
                    f"uses `{uses}` — pin to a 40-hex SHA, not a tag",
                )


# --------------------------------------------------------------------------------------------
# CROSS-WORKFLOW CHECK (check 8)
# --------------------------------------------------------------------------------------------


def _check_python_version_consistency(
    parsed: list[tuple[Path, dict[str, Any]]],
) -> Iterator[Finding]:
    """Yield a violation when workflows disagree on `env.PYTHON_VERSION`.

    Audit 2026-08-12 (oracle-vps-github-runners v2.0, Hard Rule #15): the
    interpreter is part of the gate. A patch bump is a different verdict.
    Two workflows declaring two different versions means the test suite
    is exercised on one and shipped on the other — the "passes locally"
    half of a bad CI is exactly that.

    Only workflows that DECLARE `env.PYTHON_VERSION` participate. A
    workflow that does not declare it does not contribute to the check,
    since it does not run Python (security.yml and security-deep.yml
    ship Docker images; they do not need setup-python at the workflow
    level). The check fires only when two declarations disagree.
    """
    declared: list[tuple[Path, str]] = []
    for path, doc in parsed:
        if not isinstance(doc, dict):
            continue
        env = doc.get("env")
        if not isinstance(env, dict):
            continue
        pv = env.get("PYTHON_VERSION")
        if isinstance(pv, str):
            declared.append((path, pv))
    by_value: dict[str, list[Path]] = {}
    for path, value in declared:
        by_value.setdefault(value, []).append(path)
    if len(by_value) > 1:
        # One finding that names every (workflow, value) pair so a
        # reviewer can fix the disagreement without rerunning the gate.
        pairs = ", ".join(
            f"`{p.name}` declares `{v!r}`"
            for v, paths in sorted(by_value.items())
            for p in paths
        )
        first_path = declared[0][0]
        yield (
            "error",
            "env.PYTHON_VERSION",
            f"workflows disagree on `env.PYTHON_VERSION`: {pairs}",
        )


# --------------------------------------------------------------------------------------------
# RUN-ALL-CHECKS
# --------------------------------------------------------------------------------------------


def _format_finding(path: Path, severity: str, check_name: str, location: str, message: str) -> str:
    """Render one finding to the on-disk `path:line: CHECK-NAME: location: message` shape.

    WARN findings add a `WARN: ` prefix so the existing error-line tests
    continue to find their substrings (the check name and message are
    identical to what an ERROR would have printed) and the new warn-line
    tests can grep for `WARN`. The prefix sits between the line number
    and the check name so the format stays line-anchored.
    """
    prefix = "WARN: " if severity == "warn" else ""
    return f"{path}:1: {prefix}{check_name}: {location}: {message}"


def _check_one(path: Path, doc: dict[str, Any]) -> Iterator[tuple[str, str]]:
    """Yield `(severity, formatted_line)` for every per-workflow finding on `path`.

    `doc` is the already-parsed workflow. The caller is responsible for
    reading the file and running the duplicate-key and parse-error checks
    (which short-circuit before any per-workflow check can run on a broken
    file — see `_check_one_with_parse`).
    """
    for severity, location, message in _check_timeout_minutes(doc):
        yield severity, _format_finding(path, severity, "timeout-minutes", location, message)
    for severity, location, message in _check_service_ports(doc):
        yield severity, _format_finding(path, severity, "host-port-fix", location, message)
    for severity, location, message in _check_uses_pinned(doc):
        yield severity, _format_finding(path, severity, "uses-not-pinned", location, message)
    for severity, location, message in _check_docker_preflight(doc):
        yield severity, _format_finding(path, severity, "docker-preflight", location, message)
    for severity, location, message in _check_concurrency(doc):
        yield severity, _format_finding(path, severity, "concurrency", location, message)
    for severity, location, message in _check_permissions(doc):
        yield severity, _format_finding(path, severity, "permissions", location, message)
    for severity, location, message in _check_concurrency_group_uniqueness(doc):
        yield severity, _format_finding(path, severity, "concurrency-group", location, message)


def _check_one_with_parse(path: Path) -> tuple[dict[str, Any] | None, list[tuple[str, str]]]:
    """Read `path`, run parse-time checks, return `(doc, findings)`.

    `doc` is the parsed workflow, or None if the file failed to parse or
    had duplicate keys. `findings` is a list of `(severity, formatted_line)`
    pairs for this path's parse-time findings (the per-workflow findings
    are emitted separately by `_check_one` once the caller has the doc).

    Splitting the parse from the per-workflow checks is what lets
    `main()` collect the parsed `doc` for the cross-workflow check
    without re-reading any file.
    """
    findings: list[tuple[str, str]] = []

    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        findings.append(("error", f"{path}:0: read-error: {exc}"))
        return None, findings

    # Check 1: duplicate keys. Line-based scanner — handles PyYAML's C and
    # Python builds identically, and is the only way to surface duplicates
    # that PyYAML's safe_load silently drops.
    duplicates = _find_duplicate_keys(text)
    if duplicates:
        for line, key in duplicates:
            findings.append(
                (
                    "error",
                    f"{path}:{line}: duplicate-key: key `{key}` appears more "
                    f"than once at the same indentation level — GitHub rejects "
                    f"this with `startup_failure` and the check disappears "
                    f"from the PR rollup",
                )
            )
        # Short-circuit: per-workflow checks on a file with duplicates
        # would see safe_load pick one of the colliding values
        # arbitrarily, and any finding from there would be untrustworthy.
        return None, findings

    try:
        doc = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        findings.append(("error", f"{path}:0: parse-error: {exc}"))
        return None, findings

    if not isinstance(doc, dict):
        findings.append(("error", f"{path}:0: shape-error: top-level YAML must be a mapping (GitHub Actions workflow)"))
        return None, findings

    return doc, findings


def _iter_workflows(root: Path) -> Iterator[Path]:
    if not root.is_dir():
        return
    for path in sorted(root.glob("*.yml")):
        if path.is_file():
            yield path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        default=str(WORKFLOWS_DIR),
        help="directory containing the workflow files (default: .github/workflows)",
    )
    args = parser.parse_args(argv)

    root = Path(args.root)
    if not root.is_dir():
        print(f"FAIL  workflows directory does not exist: {root}", file=sys.stderr)
        return 1

    violations = 0
    warnings = 0
    parsed: list[tuple[Path, dict[str, Any]]] = []

    for path in _iter_workflows(root):
        doc, parse_findings = _check_one_with_parse(path)
        for severity, line in parse_findings:
            print(line, file=sys.stderr)
            if severity == "warn":
                warnings += 1
            else:
                violations += 1
        if doc is not None:
            for severity, line in _check_one(path, doc):
                print(line, file=sys.stderr)
                if severity == "warn":
                    warnings += 1
                else:
                    violations += 1
            parsed.append((path, doc))

    # Cross-workflow checks (env.PYTHON_VERSION consistency) only see
    # workflows that parsed cleanly. A parse-broken file already
    # produced an error above; an absent doc is not a reason to skip the
    # cross-workflow check on the rest.
    if len(parsed) >= 2:
        for severity, location, message in _check_python_version_consistency(parsed):
            # All participating workflows are named in the message; the
            # `path` here is just the surface the on-disk line is
            # anchored to (the line number is meaningless for a
            # cross-workflow finding). Use the first parsed path as
            # the anchor.
            anchor = parsed[0][0]
            line = _format_finding(anchor, severity, "python-version-consistency", location, message)
            print(line, file=sys.stderr)
            if severity == "warn":
                warnings += 1
            else:
                violations += 1

    if violations == 0:
        files = sum(1 for _ in _iter_workflows(root))
        if warnings:
            print(f"OK    {files} workflow file(s) clean ({warnings} warning(s))")
        else:
            print(f"OK    {files} workflow file(s) clean")
        return 0

    print(
        f"FAIL  {violations} violation(s) across workflow files",
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
