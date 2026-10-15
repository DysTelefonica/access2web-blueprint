#!/usr/bin/env python3
# HARNESS-PROVENANCE: access2web-blueprint issue #854 — scripts/check_pr_size.py
"""PR size gate: additions + deletions against the 400-line review budget.

Two modes:

    pr-event <event.json>      additions, deletions and labels from a GitHub
                               ``pull_request`` webhook payload
    local [<merge-base>]       ``git diff --shortstat <merge-base>...HEAD``
                               (default merge-base: ``origin/main``)

With no subcommand the script runs ``local origin/main``, so the current CI step
keeps working until PR-C (issue #857) switches it to ``pr-event``.

The only override is the ``size:exception`` label, applied by the maintainer.
The PR body is not read.

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
from collections.abc import Iterable
from pathlib import Path

MAX_CHANGED_LINES = 400
SIZE_EXCEPTION_LABEL = "size:exception"


def has_size_exception_label(labels: Iterable[str]) -> bool:
    """True when the labels carry the override label."""
    return SIZE_EXCEPTION_LABEL in labels


def compute_local_total(cwd: Path, merge_base: str) -> int:
    """Insertions + deletions of HEAD against ``merge_base``, as GitHub counts them (rule 4)."""
    shortstat = subprocess.run(
        ["git", "diff", "--shortstat", f"{merge_base}...HEAD"],
        cwd=str(cwd),
        capture_output=True,
        text=True,
        check=True,
        encoding="utf-8",
    ).stdout
    return sum(int(count) for count in re.findall(r"(\d+) (?:insertion|deletion)", shortstat))


def _pr_event_total_and_labels(event_path: Path) -> tuple[int, list[str]]:
    pull_request = json.loads(event_path.read_text(encoding="utf-8")).get("pull_request") or {}
    additions = pull_request.get("additions")
    deletions = pull_request.get("deletions")
    if not isinstance(additions, int) or not isinstance(deletions, int):
        raise ValueError(f"{event_path}: pull_request.additions/deletions must be integers")
    labels = [label["name"] for label in pull_request.get("labels") or []]
    return additions + deletions, labels


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

    local = subparsers.add_parser("local", help="git diff --shortstat against <merge-base>")
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
    args = _build_parser().parse_args(argv)
    if args.mode is None:  # no subcommand: keep the existing CI step working
        args.mode, args.merge_base = "local", "origin/main"

    if args.mode == "local":
        total, labels = compute_local_total(Path.cwd(), args.merge_base), []
    else:
        total, labels = _pr_event_total_and_labels(args.event_json)

    if total <= MAX_CHANGED_LINES or has_size_exception_label(labels):
        return 0

    print(f"PR size: {total} (limit {MAX_CHANGED_LINES}).", file=sys.stderr)
    if args.mode == "pr-event":
        print(
            "Split the PR. Only the maintainer applies "
            f"'{SIZE_EXCEPTION_LABEL}', after the 4 questions in the PR body.",
            file=sys.stderr,
        )
    else:
        print("local mode has no override; reduce the diff or split the PR.", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
