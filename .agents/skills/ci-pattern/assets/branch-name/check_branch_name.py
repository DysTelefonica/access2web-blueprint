#!/usr/bin/env python3
# ci-pattern asset — branch name gate (P01; DysTelefonica/team-skills#806)
"""Branch-name gate: the head ref of a PR matches the convention (parameter P01).

The pattern is DATA (``--pattern``, the consumer's ``P01_branch_name_pattern``); the
default is the gentle-ai convention ``<type>/<slug>`` with no issue number. The whole
name has to match. Exempt without configuration: the default branch (``main``),
``archive/*``, catalog propagation branches (``skill-fleet/<name>``) and Dependabot's
own branches, the latter only when ``--actor`` is ``dependabot[bot]`` (the name alone is
not proof of who pushed it). Branches that pre-date the convention are listed with
``--allow`` (repeatable), never by loosening the pattern.

Exit codes: 0 conforming or exempt · 1 the name does not comply · 2 invalid input
(empty or missing ref, unreadable pattern): a gate that could not measure never says OK.
"""

from __future__ import annotations

import argparse
import re
import sys

DEFAULT_PATTERN = "^(feat|fix|chore|docs|style|refactor|perf|test|build|ci|revert)/[a-z0-9._-]+$"
ALWAYS_EXEMPT = re.compile(r"main|archive/.+|skill-fleet/[A-Za-z0-9._-]+")
DEPENDABOT = re.compile(r"dependabot/(?:pip|uv|npm_and_yarn|github_actions)/[A-Za-z0-9._/-]+")
DEPENDABOT_ACTOR = "dependabot[bot]"
EXIT_FINDING, EXIT_INVALID = 1, 2


class InvalidInputError(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


class EmptyRefError(InvalidInputError):
    def __init__(self) -> None:
        super().__init__("the head ref is empty: there is nothing to measure")


class BadPatternError(InvalidInputError):
    def __init__(self, pattern: str, exc: re.error) -> None:
        super().__init__(f"unreadable --pattern {pattern!r}: {exc}")


def compile_pattern(pattern: str) -> re.Pattern[str]:
    try:
        return re.compile(pattern)
    except re.error as exc:
        raise BadPatternError(pattern, exc) from exc


def evaluate(ref: str, pattern: re.Pattern[str], allow: frozenset[str], actor: str) -> tuple[bool, str]:
    """``(ok, note)``: the verdict for ``ref`` and why it passed or failed."""
    if not ref:
        raise EmptyRefError
    if actor == DEPENDABOT_ACTOR and DEPENDABOT.fullmatch(ref):
        return True, f"{ref}: trusted Dependabot branch"
    if ALWAYS_EXEMPT.fullmatch(ref):
        return True, f"{ref}: exempt (default branch, archive/ or skill-fleet/)"
    if ref in allow:
        return True, f"{ref}: grandfathered via --allow"
    if pattern.fullmatch(ref):
        return True, ""
    return False, (f"{ref}: branch name must match {pattern.pattern} — shape <type>/<slug>, "
                   "no issue number")


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: the output must not depend on the locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("ref", help="head ref of the PR (github.head_ref)")
    parser.add_argument("--pattern", default=DEFAULT_PATTERN, help="regex of the whole name (P01)")
    parser.add_argument("--allow", action="append", default=[], help="grandfathered branch (repeatable)")
    parser.add_argument("--actor", default="", help="who opened the PR (github.actor)")
    args = parser.parse_args(argv)
    try:
        ok, note = evaluate(args.ref, compile_pattern(args.pattern), frozenset(args.allow), args.actor)
    except InvalidInputError as exc:
        print(f"branch-name: {exc}", file=sys.stderr)
        return EXIT_INVALID
    if not ok:
        print(f"FAIL  {note}")
        return EXIT_FINDING
    print(f"OK    {note or args.ref + ' matches the pattern'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
