#!/usr/bin/env python3
# ci-pattern asset — runbook premises gate (HR-42, DysTelefonica/team-skills#167)
"""Runbook premises gate: every state a runbook assumes is seeded or
verified fail-loud (HR-42).

Scans runbook markdown under ``--root`` (recursively, ``*.md``). A
``premise`` is a prose line (outside fenced code blocks) matching any of
the versioned premise patterns — case-insensitive. Each premise MUST
carry a resource token (a single-quoted string naming the external state,
e.g. ``'e2e@apap.local'``) and the runbook MUST contain at least one line
in a fenced code block referencing that same token — the seed or
fail-loud verification command written in the runbook itself. A premise
without a resource token or without its command is a finding: prose like
«is confirmed seeded» without a command behind it is a runbook defect,
not production state (HR-42, HR-13).

The premise patterns are versioned data (``--patterns`` JSON:
``{"premise_patterns": [...]}``); the rule names no tools, jobs or
resources.

Exit codes: 0 clean · 1 findings · 2 invalid input (missing root,
missing or malformed patterns, empty subject — a root with no runbook
markdown files)."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

FENCE = re.compile(r"^\s*(```|~~~)")
SINGLE_QUOTED = re.compile(r"'([^']+)'")
PREMISE_KEYS = frozenset({"premise_patterns"})


class GateError(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


def load_patterns(path: str) -> list[str]:
    try:
        doc = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise GateError(f"patterns {path!r} unreadable: {exc}") from exc
    if not isinstance(doc, dict) or set(doc) != PREMISE_KEYS:
        raise GateError(f"patterns must carry exactly {sorted(PREMISE_KEYS)}")
    patterns = doc["premise_patterns"]
    if not isinstance(patterns, list) or not patterns \
            or not all(isinstance(p, str) and p.strip() for p in patterns):
        raise GateError('patterns "premise_patterns" must be a non-empty list '
                        "of non-empty strings")
    return patterns


def runbook_files(root: Path) -> list[Path]:
    files = [p for p in sorted(root.rglob("*.md")) if p.is_file()]
    if not files:
        raise GateError(f"no *.md runbook files under {str(root)!r} (empty subject)")
    return files


def scan_runbook(path: Path, rel: str, patterns: list[str], report: list) -> None:
    text = path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()
    in_fence = False
    premises = []
    code_lines = []
    for number, line in enumerate(lines, 1):
        if FENCE.match(line):
            in_fence = not in_fence
            continue
        (code_lines if in_fence else premises).append((number, line))
    # Solo líneas EJECUTABLES cuentan como comando: sin comentarios (#, //)
    # ni líneas vacías — un comentario «asume presente» no verifica nada.
    executable = [l for _, l in code_lines
                  if l.strip() and not l.lstrip().startswith(("#", "//"))]
    code_text = "\n".join(executable).lower()
    for number, line in premises:
        if not any(p.lower() in line.lower() for p in patterns):
            continue
        tokens = SINGLE_QUOTED.findall(line)
        if not tokens:
            report.append(f"HR-42 {rel}:{number}: state premise without a resource "
                          "token — quote the external state (e.g. 'e2e@apap.local') "
                          "so its verification command can be traced")
            continue
        if not any(tok.lower() in code_text for tok in tokens):
            report.append(f"HR-42 {rel}:{number}: state premise without a seed or "
                          f"verification command in the runbook (resource "
                          f"{tokens[0]!r} never appears in an executable code "
                          "line) — prose like «is confirmed seeded» is a runbook "
                          "defect, not production state")


def main(argv):
    ap = argparse.ArgumentParser(description="Runbook premises gate (HR-42).")
    ap.add_argument("--root", required=True)
    ap.add_argument("--patterns", required=True)
    args = ap.parse_args(argv)
    try:
        root = Path(args.root)
        if not root.is_dir():
            raise GateError(f"root {str(root)!r} is not a directory (empty subject)")
        patterns = load_patterns(args.patterns)
        files = runbook_files(root)
    except GateError as exc:
        print(f"RUNBOOK-PREMISES ERROR: {exc}", file=sys.stderr)
        return 2
    report = []
    for path in files:
        scan_runbook(path, path.relative_to(root).as_posix(), patterns, report)
    if report:
        print(f"HR-42 FAIL: {len(report)} finding(s) — runbook premises without "
              "seed or verification commands")
        for line in report:
            print(f"  - {line}")
        return 1
    print(f"RUNBOOK-PREMISES OK: {len(files)} runbook file(s) audited — every "
          "state premise carries its seed or verification command.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
