#!/usr/bin/env python3
# HARNESS-PROVENANCE: access2web-blueprint issue #875 — scripts/check_deadcode.py
"""Dead-code ratchet: vulture findings vs. .deadcode-baseline.txt.

Runs ``vulture <app_dir> --min-confidence 80`` and compares the findings to the
baseline. The ratchet only goes down: a finding that is in the baseline but no
longer reported is a failure (the maintainer must remove the entry), and a
finding not in the baseline is also a failure (the maintainer must add it).

Usage:
    check_deadcode.py [<app_dir> [<baseline>]]

Defaults: app_dir = ``app/``, baseline = ``.deadcode-baseline.txt``.

Exit codes:
    0  baseline matches vulture's findings
    1  new dead code, stale baseline entry, or vulture missing/error
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

MIN_CONFIDENCE = 80


def parse_vulture_output(text: str) -> set[str]:
    """One trimmed non-blank line per finding. ``path:line: message``."""
    return {line.strip() for line in text.splitlines() if line.strip()}


def run_vulture(app_dir: Path, vulture_bin: str = "vulture") -> set[str]:
    """Run vulture and return the set of findings. Raise FileNotFoundError if missing."""
    if shutil.which(vulture_bin) is None:
        raise FileNotFoundError(f"{vulture_bin} not in PATH")
    proc = subprocess.run(
        [vulture_bin, str(app_dir), "--min-confidence", str(MIN_CONFIDENCE)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    # vulture exits non-zero when it reports findings; the stdout is what we read.
    return parse_vulture_output(proc.stdout)


def load_baseline(path: Path) -> set[str]:
    """Read the baseline file; missing file is an empty baseline."""
    if not path.exists():
        return set()
    return parse_vulture_output(path.read_text(encoding="utf-8"))


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("app_dir", nargs="?", default="app/", type=Path)
    parser.add_argument("baseline", nargs="?", default=".deadcode-baseline.txt", type=Path)
    parser.add_argument("--vulture-bin", default="vulture")
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
    try:
        findings = run_vulture(args.app_dir, args.vulture_bin)
    except FileNotFoundError as exc:
        print(f"deadcode-ratchet: {exc}", file=sys.stderr)
        return 1
    baseline = load_baseline(args.baseline)
    new = sorted(findings - baseline)
    stale = sorted(baseline - findings)
    if new:
        print(
            f"deadcode-ratchet: {len(new)} new finding(s) outside the baseline:",
            file=sys.stderr,
        )
        for f in new:
            print(f"  {f}", file=sys.stderr)
    if stale:
        print(
            f"deadcode-ratchet: {len(stale)} stale baseline entry/ies (ratchet only goes down):",
            file=sys.stderr,
        )
        for s in stale:
            print(f"  {s}", file=sys.stderr)
    return 1 if new or stale else 0


if __name__ == "__main__":
    raise SystemExit(main())
