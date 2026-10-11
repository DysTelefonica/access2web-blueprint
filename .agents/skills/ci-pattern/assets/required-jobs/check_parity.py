#!/usr/bin/env python3
# ci-pattern asset — three-way parity check for the required-jobs aggregator (DysTelefonica/team-skills#161)
"""Parity check between a workflow, its aggregator job and the policy (HR-29).

Three sets must be equal: the workflow's jobs minus the aggregator, the
aggregator's ``needs`` and the policy's ``required_jobs``. A job wired into
``needs`` and unknown to the policy, or added to the workflow and never wired,
sits outside every verdict; this check is what notices.

The aggregator is named by ``--aggregator``, never guessed: other jobs may
declare ``needs`` of their own.

The workflow is read with a deliberately small YAML subset, stdlib only:
block-style ``jobs:`` with one plain key per job, and the aggregator's
``needs`` as a scalar (``needs: a``), a one-line flow sequence
(``needs: [a, b]``) or a block sequence (``needs:`` followed by ``- a``).
Fail-closed contract: anything outside that subset — an expression, an anchor
or alias, a merge key, an inline job, a missing ``jobs:`` block, a missing
aggregator — exits non-zero naming the construct. It never guesses a job set.

Exit codes:
    0  the three sets are equal
    1  fail-closed: a mismatch, an unreadable input, or an unsupported construct
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import NamedTuple

# Sibling module: the policy is validated by the gate's own loader, so the
# parity check and the gate can never disagree on what a valid policy is.
# Installed without the sibling gate, the parity check must still fail
# closed with an explicit message instead of a traceback.
try:
    from check_required_jobs import GateError, load_policy
except ImportError:  # the sibling gate is not installed next to this script
    class GateError(Exception):
        """Local stand-in for the sibling gate's error type."""

    def load_policy(path: str) -> dict:
        raise GateError(
            "check_required_jobs.py must be installed next to check_parity.py: "
            "the parity gate validates the policy with the gate's own loader")

JOB_ID = re.compile(r"[A-Za-z_][A-Za-z0-9_-]*")
JOB_KEY = re.compile(r"""(['"]?)([A-Za-z_][A-Za-z0-9_-]*)\1\s*:""")
NEEDS_KEY = re.compile(r"""(['"]?)needs\1\s*:(?:\s+(.*))?""")


class Line(NamedTuple):
    """One meaningful line of the workflow: comments and blanks are dropped."""

    number: int
    indent: int
    text: str
    tabbed: bool  # the indentation continues with a tab after the spaces


def strip_comment(text: str) -> str:
    """Drop a trailing ``# comment``; a ``#`` inside quotes is content."""
    quote = None
    for index, char in enumerate(text):
        if quote:
            if char == quote:
                quote = None
        elif char in "'\"":
            quote = char
        elif char == "#" and (index == 0 or text[index - 1] in " \t"):
            return text[:index].rstrip()
    return text.rstrip()


def read_lines(text: str) -> list[Line]:
    lines: list[Line] = []
    for number, raw in enumerate(text.splitlines(), start=1):
        body = raw.lstrip(" ")
        content = strip_comment(body.strip())
        if content:
            line = Line(number, len(raw) - len(body), content, body.startswith("\t"))
            # Fail closed on tabbed indentation for EVERY meaningful line:
            # job keys, job bodies and needs items alike. A tab hides the
            # structural level this parser relies on.
            if line.tabbed:
                raise unsupported(line, "tab in indentation")
            lines.append(line)
    return lines


def unsupported(line: Line, what: str) -> GateError:
    return GateError(f"line {line.number}: {what}: {line.text}")


def require_spaces(line: Line) -> None:
    """Structural lines must be indented with spaces; a tab hides the level."""
    if line.tabbed:
        raise unsupported(line, "tab in indentation")


def describe(value: str) -> str:
    """Name the YAML construct a non-plain value is, for the message."""
    if "${{" in value:
        return "expression"
    for prefix, name in (("*", "alias"), ("&", "anchor"), ("|", "block scalar"),
                         (">", "block scalar"), ("{", "flow mapping"),
                         ("[", "nested sequence"), ("!", "tag")):
        if value.startswith(prefix):
            return name
    return "value that is not a plain job id"


def job_id(value: str, line: Line) -> str:
    """Read one needs entry; only a plain or quoted job id is a job id."""
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "'\"":
        value = value[1:-1]
    if not JOB_ID.fullmatch(value):
        raise unsupported(line, f"unsupported needs value ({describe(value)})")
    return value


def split_jobs(lines: list[Line]) -> dict[str, list[Line]]:
    """Return each job id with the lines of its body, in workflow order."""
    starts = [i for i, line in enumerate(lines) if line.indent == 0 and line.text.startswith("jobs:")]
    for line in lines:
        if line.indent == 0:
            require_spaces(line)
    if not starts:
        raise GateError("no top-level 'jobs:' block")
    if len(starts) > 1:
        raise unsupported(lines[starts[1]], "duplicate top-level 'jobs:' block")
    if lines[starts[0]].text != "jobs:":
        raise unsupported(lines[starts[0]], "'jobs:' with an inline value")

    block: list[Line] = []
    for line in lines[starts[0] + 1:]:
        if line.indent == 0:
            break
        block.append(line)
    if not block:
        raise GateError("the 'jobs:' block defines no job")

    job_indent = block[0].indent
    jobs: dict[str, list[Line]] = {}
    current: list[Line] | None = None
    for line in block:
        if line.indent > job_indent:
            # Belongs to the job above; only the aggregator's body is read later.
            current.append(line)
            continue
        require_spaces(line)
        if line.indent < job_indent:
            raise unsupported(line, "indentation shallower than the job keys")
        match = JOB_KEY.fullmatch(line.text)
        if not match:
            # An inline mapping, an anchor, an alias or a merge key: the job's
            # definition lives somewhere this parser does not follow.
            raise unsupported(line, "unsupported job entry (expected a plain 'job-id:' line)")
        name = match.group(2)
        if name in jobs:
            raise unsupported(line, f"duplicate job '{name}'")
        current = jobs[name] = []
    return jobs


def block_items(body: list[Line], start: int, key: Line) -> list[str]:
    """Read the ``- item`` lines that follow an empty ``needs:`` key."""
    items: list[str] = []
    item_indent = None
    index = start
    require_spaces(key)
    while index < len(body):
        line = body[index]
        is_item = line.text == "-" or line.text.startswith("- ")
        if line.indent < key.indent or not is_item:
            break
        require_spaces(line)
        if item_indent is None:
            item_indent = line.indent
        elif line.indent != item_indent:
            raise unsupported(line, "unsupported needs value (unevenly indented sequence)")
        value = line.text[1:].strip()
        if not value:
            raise unsupported(line, "unsupported needs value (empty entry)")
        if not value.startswith(("'", '"')) and (value.endswith(":") or ": " in value):
            raise unsupported(line, "unsupported needs value (mapping entry)")
        items.append(job_id(value, line))
        index += 1
    if not items:
        raise unsupported(key, "empty needs (no value and no '- item' lines)")
    if index < len(body) and body[index].indent > key.indent:
        # Something deeper than the key follows the items: a continuation
        # line or a nested structure this parser would silently drop.
        raise unsupported(body[index], "unsupported needs value (content after the sequence)")
    return items


def parse_needs(body: list[Line], aggregator: str) -> list[str]:
    """Read the aggregator's ``needs`` from the lines of its body."""
    if not body:
        raise GateError(f"aggregator job '{aggregator}' declares no needs")
    key_indent = body[0].indent
    needs: list[str] | None = None
    for index, line in enumerate(body):
        if line.indent > key_indent:
            continue
        require_spaces(line)
        if line.indent < key_indent:
            raise unsupported(line, "indentation shallower than the job's keys")
        if line.text.startswith("<<"):
            # A merge key can bring a needs list from an anchor defined elsewhere.
            raise unsupported(line, "unsupported merge key in the aggregator job")
        match = NEEDS_KEY.fullmatch(line.text)
        if not match:
            continue
        if needs is not None:
            raise unsupported(line, "duplicate 'needs' key in the aggregator job")
        value = (match.group(2) or "").strip()
        if not value:
            needs = block_items(body, index + 1, line)
        elif value.startswith("["):
            if not value.endswith("]"):
                raise unsupported(line, "unsupported needs value (multi-line flow sequence)")
            entries = [entry.strip() for entry in value[1:-1].split(",")]
            if entries and not entries[-1]:
                entries.pop()  # "[a, b,]" and "[]" both end in an empty entry
            if not entries:
                raise unsupported(line, "unsupported needs value (no entries)")
            needs = [job_id(entry, line) for entry in entries]
        else:
            needs = [job_id(value, line)]
        if len(set(needs)) != len(needs):
            raise unsupported(line, "duplicate needs entry")
    if needs is None:
        raise GateError(f"aggregator job '{aggregator}' declares no needs")
    return needs


def read_workflow(path: str, aggregator: str) -> tuple[list[str], list[str]]:
    """Return the workflow's job ids and the aggregator's needs."""
    try:
        text = Path(path).read_bytes().decode("utf-8-sig")
    except OSError as exc:
        raise GateError(f"unreadable workflow file: {exc}") from exc
    except UnicodeDecodeError as exc:
        raise GateError(f"workflow file is not valid UTF-8: {exc}") from exc
    try:
        jobs = split_jobs(read_lines(text))
        if aggregator not in jobs:
            raise GateError(f"aggregator job '{aggregator}' is not defined in the workflow")
        return list(jobs), parse_needs(jobs[aggregator], aggregator)
    except GateError as exc:
        raise GateError(f"{path}: {exc}") from exc


def parity_mismatches(jobs: list[str], needs: list[str], policy_jobs: list[str],
                      aggregator: str) -> list[str]:
    """Compare jobs − {aggregator}, the aggregator's needs and the policy set."""
    workflow_jobs = set(jobs) - {aggregator}
    deps = set(needs)
    known = set(policy_jobs)
    problems: list[str] = []
    if aggregator in known:
        problems.append(f"policy lists the aggregator '{aggregator}' itself as a required job")
        known = known - {aggregator}
    for job in sorted(workflow_jobs - known):
        problems.append(f"workflow job '{job}' is not in the policy")
    for job in sorted(known - workflow_jobs):
        problems.append(f"policy job '{job}' is not defined in the workflow")
    for job in sorted(deps - known):
        problems.append(f"aggregator needs '{job}' is not in the policy")
    for job in sorted(known - deps):
        problems.append(f"policy job '{job}' is missing from the aggregator needs")
    return problems


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="fail-closed workflow/needs/policy parity check")
    parser.add_argument("--workflow", required=True, help="path to the workflow YAML file")
    parser.add_argument("--aggregator", required=True, help="job id of the aggregator job")
    parser.add_argument("--policy", required=True, help="path to the required-jobs policy JSON")
    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="backslashreplace")
    try:
        policy = load_policy(args.policy)
        jobs, needs = read_workflow(args.workflow, args.aggregator)
    except GateError as exc:
        print(f"FAIL fail-closed: {exc}", file=sys.stderr)
        return 1
    problems = parity_mismatches(jobs, needs, policy["required_jobs"], args.aggregator)
    if problems:
        print(f"PARITY: FAIL (aggregator: {args.aggregator})")
        print("\n".join(f"  {item}" for item in problems))
        return 1
    print(
        f"PARITY: OK (aggregator: {args.aggregator}; {len(needs)} jobs agree across "
        "the workflow, the aggregator needs and the policy)"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
