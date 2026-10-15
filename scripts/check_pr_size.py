#!/usr/bin/env python3
# HARNESS-PROVENANCE: access2web-blueprint issue #854 — scripts/check_pr_size.py
"""PR size gate: additions + deletions against the 400-line review budget.

Two modes:

    pr-event <event.json>      additions, deletions and labels from a GitHub
                               ``pull_request`` webhook payload
    local [<merge-base>]       ``git diff --shortstat <merge-base>...HEAD``
                               (default merge-base: ``origin/main``)

When invoked with no subcommand the script defaults to ``local origin/main``,
which keeps the existing ``python scripts/check_pr_size.py`` CI step working
unchanged (PR-C, issue #857, will switch CI to ``pr-event``).

Override is the ``size:exception`` label applied by the maintainer. The
``size-exception-reason:`` field on the PR body is no longer consulted.

Exit codes:
    0  within budget, or overridden by the ``size:exception`` label
    1  over budget, or malformed payload
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Iterable

MAX_CHANGED_LINES = 400
SIZE_EXCEPTION_LABEL = "size:exception"


def has_size_exception_label(labels: Iterable[str]) -> bool:
    """True when the iterable carries the override label."""
    return SIZE_EXCEPTION_LABEL in set(labels)


def compute_local_total(cwd: Path, merge_base: str) -> int:
    """Total insertions + deletions of HEAD vs ``merge_base`` in ``cwd``.

    Parses ``git diff --shortstat`` so it equals the API total for the same range
    (rule 4 of issue #854).
    """
    result = subprocess.run(
        ["git", "diff", "--shortstat", f"{merge_base}...HEAD"],
        cwd=str(cwd),
        capture_output=True,
        text=True,
        check=True,
        encoding="utf-8",
    )
    inserts = 0
    deletes = 0
    for token in result.stdout.split(","):
        token = token.strip()
        insertion = re.match(r"(\d+) insertion", token)
        deletion = re.match(r"(\d+) deletion", token)
        if insertion:
            inserts = int(insertion.group(1))
        elif deletion:
            deletes = int(deletion.group(1))
    return inserts + deletes


def _label_names(payload: dict) -> list[str]:
    """Label names from a ``pull_request`` payload — accepts dicts or strings."""
    pull_request = payload.get("pull_request") or {}
    labels = pull_request.get("labels") or []
    names: list[str] = []
    for label in labels:
        if isinstance(label, str):
            names.append(label)
        elif isinstance(label, dict):
            name = label.get("name")
            if isinstance(name, str):
                names.append(name)
    return names


def _pr_event_total_and_labels(event_path: Path) -> tuple[int, list[str]]:
    payload = json.loads(event_path.read_text(encoding="utf-8"))
    pull_request = payload.get("pull_request") or {}
    additions = pull_request.get("additions")
    deletions = pull_request.get("deletions")
    if not isinstance(additions, int) or not isinstance(deletions, int):
        raise ValueError(
            f"{event_path}: pull_request.additions/deletions must be integers"
        )
    return additions + deletions, _label_names(payload)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="mode")

    pr_event = subparsers.add_parser(
        "pr-event",
        help="read additions, deletions and labels from a GitHub pull_request event",
    )
    pr_event.add_argument("event_json", type=Path)

    local = subparsers.add_parser(
        "local", help="git diff --shortstat against <merge-base>"
    )
    local.add_argument(
        "merge_base",
        nargs="?",
        default="origin/main",
        help="merge-base ref (default: origin/main)",
    )
    return parser


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8 (Hard Rule 8)."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = _build_parser()
    args = parser.parse_args(argv)

    # Default to `local origin/main` so the existing CI step keeps working.
    if args.mode is None:
        args.mode = "local"
        args.merge_base = "origin/main"

    if args.mode == "local":
        total = compute_local_total(Path.cwd(), args.merge_base)
        labels: list[str] = []
    else:  # pr-event
        total, labels = _pr_event_total_and_labels(args.event_json)

    if total <= MAX_CHANGED_LINES:
        return 0
    if has_size_exception_label(labels):
        return 0

    print(f"PR size: {total} (limit {MAX_CHANGED_LINES}).", file=sys.stderr)
    if args.mode == "pr-event":
        print(
            f"Apply the '{SIZE_EXCEPTION_LABEL}' label or split the PR.",
            file=sys.stderr,
        )
    else:
        print(
            "local mode has no override; reduce the diff or split the PR.",
            file=sys.stderr,
        )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
