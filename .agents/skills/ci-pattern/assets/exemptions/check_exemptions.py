#!/usr/bin/env python3
# ci-pattern asset — exemption identity gate (HR-33, DysTelefonica/team-skills#245)
"""Exemption identity gate: every gate exemption declared as data carries a
verifiable identity (HR-33).

Scans the JSON documents given with ``--file`` — recursively, so baselines,
policies and matrices are all covered — for exemption-shaped data:

- an object with a ``relaxation`` key (ratchet baselines, required-jobs
  policies), or
- an object with an ``exemption`` key (matrix entries, gate policy files).

Every exemption MUST carry a verifiable identity inside itself: a non-empty
``actor`` string (who granted it) or a positive ``issue`` integer (where it
was granted and can be audited). A bare-string exemption, an object without
identity, or an identity that is empty or non-positive is a finding: an
exemption nobody can be held accountable for is not governance (HR-33).

Documents that declare no exemptions are clean. Malformed JSON, a
non-object document, a missing file, or an empty subject (no files) exit 2
— the gate never says OK about input it could not read (HR-3).

Exit codes: 0 clean · 1 findings · 2 invalid input or empty subject.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

EXEMPTION_KEYS = ("relaxation", "exemption")


class GateError(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


def has_identity(exemption) -> bool:
    if not isinstance(exemption, dict):
        return False
    actor = exemption.get("actor")
    issue = exemption.get("issue")
    if isinstance(actor, str) and actor.strip():
        return True
    if isinstance(issue, int) and not isinstance(issue, bool) and issue > 0:
        return True
    return False


def walk(node, path, found):
    if isinstance(node, dict):
        for key in EXEMPTION_KEYS:
            if key in node:
                found.append((path + [key], node[key]))
        for key, value in node.items():
            walk(value, path + [str(key)], found)
    elif isinstance(node, list):
        for i, value in enumerate(node):
            walk(value, path + [f"[{i}]"], found)


def load_document(path: str) -> object:
    try:
        doc = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise GateError(f"document {path!r} unreadable: {exc}") from exc
    if not isinstance(doc, dict):
        raise GateError(f"document {path!r} must be a JSON object")
    return doc


def main(argv):
    ap = argparse.ArgumentParser(description="Exemption identity gate (HR-33).")
    ap.add_argument("--file", action="append", default=[],
                    help="JSON document to scan for exemptions (repeatable)")
    args = ap.parse_args(argv)
    if not args.file:
        print("EXEMPTIONS ERROR: no documents to scan (empty subject)", file=sys.stderr)
        return 2
    report = []
    try:
        for path in args.file:
            doc = load_document(path)
            locations = []
            walk(doc, [], locations)
            for location, exemption in locations:
                if has_identity(exemption):
                    continue
                where = f"{path}:{'.'.join(location)}"
                detail = ("a bare string carries no identity" if isinstance(exemption, str)
                          else "identity must be a non-empty \"actor\" or a positive \"issue\"")
                report.append(f"{where}: exemption without verifiable identity — {detail}")
    except GateError as exc:
        print(f"EXEMPTIONS ERROR: {exc}", file=sys.stderr)
        return 2
    if report:
        print(f"HR-33 FAIL: {len(report)} exemption(s) without verifiable identity")
        for line in report:
            print(f"  - {line}")
        return 1
    print("EXEMPTIONS OK: every declared exemption carries a verifiable identity.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
