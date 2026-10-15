#!/usr/bin/env python3
"""Fail closed unless every mandatory suite collects at least one test.

Issue #882 (PR-1): a green CI with zero tests in a suite is a false guarantee.
The guard runs ``pytest --collect-only -q -m <expr>`` for each mandatory suite
and fails closed if any suite collects zero tests. Renaming a marker, deleting
a directory, or skipping a fixture silently now fails this gate — exactly the
class of drift this script exists to catch.

Three suites are pinned:

    unit         — tests without `integration` or `e2e_smoke` markers
    integration  — @pytest.mark.integration (Postgres + out-of-process deps)
    e2e_smoke    — @pytest.mark.e2e_smoke (gate-smoke subprocess tests)

The `unit` suite is by exclusion (``-m "not integration and not e2e_smoke"``),
not by an explicit marker. This is deliberate: marking every unit test with
``@pytest.mark.unit`` is a much larger surface, and the exclusion gives the same
coverage with no extra annotation burden.

Exit codes:
    0  every suite collected at least one test
    1  any suite collected zero tests, OR pytest itself failed

The script is stdlib only; ``pytest`` is the only external dependency, and it
is already on every CI runner.
"""

from __future__ import annotations

import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Suite:
    """One mandatory suite the guard pins.

    `marker_expr` is passed to ``pytest -m``. It must be valid against the
    markers declared in ``app/pyproject.toml``; with ``--strict-markers`` on,
    an unregistered marker fails pytest at COLLECTION time and is caught as
    a guard violation rather than slipping through silently.
    """

    name: str
    marker_expr: str


# The three suites the gate pins. Order matters only for the error message
# (it lists failures in declaration order). Keep this tuple the single
# source of truth: anything that wants to know which suites are mandatory
# reads it from here.
SUITES: tuple[Suite, ...] = (
    Suite(name="unit", marker_expr="not integration and not e2e_smoke"),
    Suite(name="integration", marker_expr="integration"),
    Suite(name="e2e_smoke", marker_expr="e2e_smoke"),
)


# Captures the X in "X/Y tests collected". pytest prints this line in the
# multi-line summary; we accept either the "X/Y" form (when other markers
# are also being collected) or the bare "X tests collected" form (when
# one suite is collected in isolation). Either way the first integer is
# the count for the expression in question.
_COUNT_RE = re.compile(rb"(\d+)(?:/\d+)? tests? collected")


def _count(repo_root: Path, marker_expr: str) -> tuple[int, str]:
    """Return ``(count, stderr)`` from a pytest --collect-only run."""
    proc = subprocess.run(
        [
            "pytest",
            "-c",
            "app/pyproject.toml",
            "--rootdir",
            "app",
            "--collect-only",
            "-q",
            "-m",
            marker_expr,
        ],
        capture_output=True,
        cwd=str(repo_root),
        check=False,
    )
    stderr = proc.stderr.decode("utf-8", errors="replace")
    if proc.returncode != 0:
        return -1, stderr
    match = _COUNT_RE.search(proc.stdout)
    if match is None:
        # pytest with no matching tests prints a "no tests collected" message
        # rather than the "X tests collected" line we match on. Treat that
        # as zero so the guard's contract is the same either way.
        return 0, stderr
    return int(match.group(1)), stderr


def check_suites(repo_root: Path) -> list[str]:
    """Return the list of suite-level violations found in the repo.

    A violation is a message describing one of the three ways a suite can
    fail the guard contract:

        1. The suite collected zero tests. This is the case the guard
           exists to catch: a marker rename, a directory move, a typo.
        2. pytest itself failed during collection (e.g. an unparseable
           test file, an unregistered marker under ``--strict-markers``).

    The list is empty when every suite collected at least one test and
    pytest ran cleanly.
    """
    violations: list[str] = []

    for suite in SUITES:
        count, stderr = _count(repo_root, suite.marker_expr)
        if count < 0:
            violations.append(
                f"{suite.name}: pytest collection failed; the marker "
                f"expression {suite.marker_expr!r} probably references a "
                f"marker that is not registered in app/pyproject.toml. "
                f"Re-run `pytest --collect-only -m {suite.marker_expr}` "
                f"locally to see the full error:\n{stderr.strip()}"
            )
            continue
        if count == 0:
            violations.append(
                f"{suite.name}: collected 0 tests under "
                f"`-m {suite.marker_expr}`. Either (a) a test directory "
                f"was renamed, (b) a marker was renamed or dropped, or "
                f"(c) no test in the suite carries the right marker. "
                f"Add a test with the right marker, or if this suite is "
                f"no longer mandatory, edit SUITES in "
                f"`scripts/check_suites_populated.py`."
            )

    return violations


def _pin_output_encoding() -> None:
    """Make gate output independent of the runner locale."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    repo_root = Path(argv[0]).resolve() if argv else Path.cwd()
    violations = check_suites(repo_root)
    if violations:
        for violation in violations:
            print(f"FAIL suites guard: {violation}", file=sys.stderr)
        return 1
    print("suites guard: every mandatory suite collected at least one test")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
