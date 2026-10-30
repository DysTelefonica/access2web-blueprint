#!/usr/bin/env python3
"""Docs-only lane classifier: is this diff documentation and nothing else?

Issue #1368: a pull request whose whole diff lives under ``docs/`` does not need
the heavy suite. The ``pr-gate`` job publishes this verdict as its
``docs_changed`` output, and the required-jobs policy accepts the heavy jobs'
skips only when that marker is exactly ``true``
(``.agents/skills/ci-pattern/assets/required-jobs/README.md``, ``docs_only_lane``).

The classifier is fail-closed. It answers ``true`` only for a NON-EMPTY changed
set whose every path is documentation: one single path outside ``docs/`` pays the
whole heavy suite, and an empty diff is a known no-op, not a documentation diff.
The anti-self-exemption toll is implicit — this script, the workflow and the
policy live outside ``docs/``, so editing them never exempts anything.

Usage:
    check_docs_only_lane.py

Reads the changed paths from stdin, one per line, and prints ``true`` or
``false``.

Exit codes:
    0  the verdict was printed
    1  stdin could not be read
"""

from __future__ import annotations

import sys
from collections.abc import Iterable

#: The documentation surface: everything under ``docs/``.
DOCS_PREFIX = "docs/"


def is_docs_path(path: str) -> bool:
    """True when ``path`` belongs to the documentation surface."""
    return path.startswith(DOCS_PREFIX)


def docs_only(paths: Iterable[str]) -> bool:
    """True only when the changed set is non-empty and every path is documentation.

    An empty set is a no-op, not a documentation diff, so it answers ``False``:
    the heavy suite runs and the question never has to be asked.
    """
    collected = list(paths)
    if not collected:
        return False
    return all(is_docs_path(path) for path in collected)


def read_paths(stream: Iterable[str]) -> list[str]:
    """The changed paths of ``stream``, one per line, without the blank ones."""
    return [line.strip() for line in stream if line.strip()]


def main() -> int:
    try:
        paths = read_paths(sys.stdin)
    except OSError as exc:  # unreadable stdin is a doubt, never a documentation diff
        print(f"check_docs_only_lane: cannot read the changed paths: {exc}", file=sys.stderr)
        return 1
    print("true" if docs_only(paths) else "false")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
