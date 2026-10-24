#!/usr/bin/env python3
"""Classify which CI surface a pull request's diff touches (issue #1376).

The heavy jobs of `ci.yml` only look at the code they analyze: `e2e-smoke`
builds the image of the app and runs the Playwright specs of `e2e/`, and
`codeql` scans the Python of the repository. A pull request that does not
touch those paths pays for the whole job for nothing — 153 pull request runs
measured 58 chores, 47 CI changes and 16 documentation edits against 23 fixes
and 7 features, so roughly four out of five pull requests do not touch the
application at all.

    printf '%s\n' "$CHANGED" | python scripts/check_change_scope.py e2e

The paths of the diff arrive on stdin, one per line. The answer is `true`
when the diff touches the surface and `false` when it does not. The
fail-closed half lives in the caller: the `pr-gate` step starts every marker
at `true` and only lowers it with the complete list of paths, so a diff that
could not be computed leaves the heavy jobs running.
"""

from __future__ import annotations

import sys
from collections.abc import Iterable

#: The surface each heavy job analyzes: directory prefixes plus exact files.
SURFACES: dict[str, tuple[tuple[str, ...], tuple[str, ...]]] = {
    # `e2e-smoke` builds the image of the app (`app/`, with its Dockerfile and
    # its migrations) and runs the specs of `e2e/`.
    "e2e": (("app/", "e2e/"), ()),
    # `codeql` scans the Python of the repository: the app and the scripts.
    "codeql": (("app/", "scripts/"), ()),
}


def touches_surface(paths: Iterable[str], surface: str) -> bool:
    """True when any path of the diff belongs to the surface."""
    prefixes, files = SURFACES[surface]
    for path in paths:
        if path in files or any(path.startswith(prefix) for prefix in prefixes):
            return True
    return False


def main(argv: list[str]) -> int:
    if len(argv) != 2 or argv[1] not in SURFACES:
        surfaces = "|".join(sorted(SURFACES))
        print(f"usage: {argv[0]} <{surfaces}>", file=sys.stderr)
        return 2
    paths = [line.strip() for line in sys.stdin.read().splitlines() if line.strip()]
    print("true" if touches_surface(paths, argv[1]) else "false")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
