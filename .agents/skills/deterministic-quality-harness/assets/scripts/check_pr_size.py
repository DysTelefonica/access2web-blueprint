#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v1.9 — assets/scripts/check_pr_size.py
"""PR size gate: a review budget, enforced.

Counts the added plus deleted lines of the PR between the merge base and HEAD, as GitHub reports
them. There are no exclusions: lockfiles and generated files count (A1, #518).

The only override is the ``size:exception`` label, applied by a separate identity (HR-33, #430).
The CI passes the GitHub login that applied the label as ``SIZE_EXCEPTION_APPROVER``; that login is
recorded only when the label is applied. An absent approver is treated as the PR author
(self-approval) and fails closed. The PR body carries no override field: the maintainer applies the
label only after the agent answers the four questions in the PR body (A1).

Exit codes:
    0  within budget, or over budget with the label applied by a separate identity
    1  over budget without a valid override, or the label applied by the PR author
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------

MAX_CHANGED_LINES = 400

OVERRIDE_MARKER = "size:exception"

# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------


def _git(*args: str) -> str:
    result = subprocess.run(
        ["git", *args], capture_output=True, text=True, check=False, encoding="utf-8"
    )
    if result.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout


def changed_lines(base_ref: str) -> tuple[int, list[tuple[str, int]]]:
    """Return the total counted lines and the per-file breakdown, largest first."""
    raw = _git("diff", "--numstat", f"{base_ref}...HEAD")
    per_file: list[tuple[str, int]] = []
    total = 0
    for line in raw.splitlines():
        parts = line.split("\t")
        if len(parts) != 3:
            continue
        added, deleted, path = parts
        if added == "-" or deleted == "-":  # binary file, not reviewable line by line
            continue
        count = int(added) + int(deleted)
        total += count
        per_file.append((path, count))
    per_file.sort(key=lambda item: item[1], reverse=True)
    return total, per_file


def override_has_separate_approver(pr_author: str, approver_login: str) -> bool:
    """Return True when the ``size:exception`` label was applied by a different actor.

    The CI passes the GitHub login that applied the label as
    ``SIZE_EXCEPTION_APPROVER``. An absent approver is treated as the PR author
    — i.e. a self-approval fails closed (#430, alcance punto 1; HR-33). When
    the PR author is itself unknown, separation cannot be proved and the gate
    also fails closed.
    """

    author = pr_author.strip().lower()
    approver = approver_login.strip().lower()
    if not author or not approver:
        return False
    return author != approver


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8.

    Python picks the output encoding from the platform locale, so the same gate emits different
    bytes on a Windows workstation (cp1252) and a Linux runner (utf-8) — and a non-encodable
    character crashes the write outright. A harness that claims determinism cannot let its own
    output depend on where it ran.
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--base-ref",
        default=os.environ.get("BASE_REF", "origin/main"),
        help="ref to diff against (default: $BASE_REF or origin/main)",
    )
    parser.add_argument("--json", action="store_true", help="emit the indicator envelope")
    args = parser.parse_args(argv)

    try:
        total, per_file = changed_lines(args.base_ref)
    except RuntimeError as exc:
        if args.json:
            print(json.dumps({"gate": "pr_size", "status": "error", "detail": str(exc)}))
        else:
            print(f"FAIL  {exc}", file=sys.stderr)
        return 1

    pr_author = os.environ.get("PR_AUTHOR", "")
    approver_login = os.environ.get("SIZE_EXCEPTION_APPROVER", "")
    has_separate_approver = override_has_separate_approver(pr_author, approver_login)

    if args.json:
        passed = total <= MAX_CHANGED_LINES or has_separate_approver
        findings: list[dict[str, str]] = []
        if not passed:
            findings.extend(
                {"file": path, "line": 0, "detail": f"{count} changed line(s)"}
                for path, count in per_file[:10]
            )
            if approver_login:
                findings.append(
                    {
                        "file": "<approver>",
                        "line": 0,
                        "detail": (
                            "size:exception was applied by the PR author; "
                            "a separate identity MUST apply it"
                        ),
                    }
                )
        print(
            json.dumps(
                {
                    "gate": "pr_size",
                    "status": "pass" if passed else "fail",
                    "indicators": {
                        "changed_lines": total,
                        "files_changed": len(per_file),
                        "overridden": total > MAX_CHANGED_LINES and has_separate_approver,
                        "override_has_separate_approver": has_separate_approver,
                    },
                    "ceilings": {"changed_lines": MAX_CHANGED_LINES},
                    "findings": findings,
                },
                indent=2,
            )
        )
        return 0 if passed else 1

    if total <= MAX_CHANGED_LINES:
        print(f"OK    {total} changed line(s), budget is {MAX_CHANGED_LINES}")
        return 0

    if has_separate_approver:
        print(
            f"OK    {total} changed line(s) over budget, overridden by '{OVERRIDE_MARKER}' "
            f"applied by '{approver_login}'"
        )
        return 0

    print(f"FAIL  {total} changed line(s), budget is {MAX_CHANGED_LINES}")
    for path, count in per_file[:10]:
        print(f"        {count:>6}  {path}")
    if not approver_login:
        print(
            f"        split the PR, or have a maintainer apply '{OVERRIDE_MARKER}' once the "
            f"PR body answers the four questions (SIZE_EXCEPTION_APPROVER must be the login "
            f"that applied it)"
        )
    else:
        print(
            f"        '{OVERRIDE_MARKER}' was applied by '{approver_login}' — the same login as "
            f"the PR author. A separate identity MUST apply it (HR-33, #430)."
        )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
