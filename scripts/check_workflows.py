#!/usr/bin/env python3
"""CI gate on .github/workflows/*.yml — issue #135.

Four checks, all from PRs #126/#129 applied by hand and never re-verified:
1. No duplicate YAML keys (GitHub rejects them with startup_failure that
   silently disappears from the rollup).
2. Every job declares timeout-minutes (GitHub default is 360 minutes — the
   exact failure mode that hit security-deep until PR #131 fixed it).
3. No service container fixes a host port (ports: - 5432:5432 silently
   shares a DB between runners; vacuous today, future-proofs the next
   person who adds a service).
4. Every uses: pinned by a 40-hex SHA (blinds against future tag drift).

Hard Rule 1: no || true, no continue-on-error. Exit 0 = clean, exit 1 = any
violation.
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

#: `ports:` value that fixes the host side, e.g. `5432:5432`. A bare
#: number or `number:` (empty host) does not.
_HOST_PORT_FIX = re.compile(r"^(?:\d{1,3}(?:\.\d{1,3}){3}|\d+):\d+$")


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
# CHECKS 2 / 3 / 4 (need a parsed dict)
# --------------------------------------------------------------------------------------------
def _check_timeout_minutes(doc: dict[str, Any]) -> Iterator[tuple[str, str]]:
    """Yield `(location, message)` for every job missing `timeout-minutes`."""
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
            f"jobs.{job_name}",
            f"job `{job_name}` does not declare `timeout-minutes`{hint}",
        )


def _check_service_ports(doc: dict[str, Any]) -> Iterator[tuple[str, str]]:
    """Yield `(location, message)` for every fixed host port in `services:`."""
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
                        f"jobs.{job_name}.services.{service_name}.ports",
                        f"service `{service_name}` in job `{job_name}` fixes host port `{entry}`",
                    )


def _check_uses_pinned(doc: dict[str, Any]) -> Iterator[tuple[str, str]]:
    """Yield `(location, message)` for every `uses:` not pinned by 40-hex SHA."""
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
                    f"jobs.{job_name}.steps",
                    f"step `{step.get('name', '?')}` in job `{job_name}` "
                    f"uses `{uses}` without `@<ref>`",
                )
                continue
            ref = uses.rsplit("@", 1)[1]
            if not _SHA_PIN.match(ref):
                yield (
                    f"jobs.{job_name}.steps",
                    f"step `{step.get('name', '?')}` in job `{job_name}` "
                    f"uses `{uses}` — pin to a 40-hex SHA, not a tag",
                )


# --------------------------------------------------------------------------------------------
# RUN-ALL-CHECKS
# --------------------------------------------------------------------------------------------


def _check_one(path: Path) -> Iterator[str]:
    """Yield `path:line: CHECK-NAME: message` strings for every violation."""
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        yield f"{path}:0: read-error: {exc}"
        return

    # Check 1: duplicate keys. Line-based scanner — handles PyYAML's C and
    # Python builds identically, and is the only way to surface duplicates
    # that PyYAML's safe_load silently drops.
    duplicates = _find_duplicate_keys(text)
    if duplicates:
        for line, key in duplicates:
            yield (
                f"{path}:{line}: duplicate-key: key `{key}` appears more "
                f"than once at the same indentation level — GitHub rejects "
                f"this with `startup_failure` and the check disappears "
                f"from the PR rollup"
            )
        return

    # For checks 2 / 3 / 4 we need a dict. safe_load dedupes (which is fine —
    # the file has no duplicates at this point) and resolves aliases/merges.
    try:
        doc = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        yield f"{path}:0: parse-error: {exc}"
        return

    # For checks 2 / 3 / 4 we need a dict. safe_load dedupes (which is fine —
    # the file has no duplicates at this point) and resolves aliases/merges.
    try:
        doc = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        yield f"{path}:0: parse-error: {exc}"
        return

    if not isinstance(doc, dict):
        yield (f"{path}:0: shape-error: top-level YAML must be a mapping (GitHub Actions workflow)")
        return

    for location, message in _check_timeout_minutes(doc):
        yield f"{path}:1: timeout-minutes: {location}: {message}"
    for location, message in _check_service_ports(doc):
        yield f"{path}:1: host-port-fix: {location}: {message}"
    for location, message in _check_uses_pinned(doc):
        yield f"{path}:1: uses-not-pinned: {location}: {message}"


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
    for path in _iter_workflows(root):
        for line in _check_one(path):
            print(line, file=sys.stderr)
            violations += 1

    if violations == 0:
        files = sum(1 for _ in _iter_workflows(root))
        print(f"OK    {files} workflow file(s) clean")
        return 0

    print(
        f"FAIL  {violations} violation(s) across workflow files",
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
