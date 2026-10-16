#!/usr/bin/env python3
# ci-pattern asset — vulture -> ratchet normalizer (DysTelefonica/team-skills#528)
"""Normalize `vulture --format json` output into the ratchet's input shape.

The ratchet (`check_ratchet.py`) knows no measuring tool: it consumes a
list of findings with the stable identity ``(rule, path, fingerprint)``.
vulture emits per-finding objects with ``file``, ``type``, ``message``,
``first_line`` and (for symbols) ``function`` or ``class``. This script
maps each ``type`` to a stable ``rule`` and computes ``fingerprint`` as
the SHA-1 of ``rule|path|symbol_name`` (content- and route-based, never
the mutable line number — HR-16).

Recognized vulture ``type`` values:

- ``unused-function`` -> ``rule="dead-code:unused-function"``
- ``unused-variable`` -> ``rule="dead-code:unused-variable"``
- ``unused-class`` -> ``rule="dead-code:unused-class"``
- ``unused-attribute`` -> ``rule="dead-code:unused-attribute"``
- ``unused-import`` -> ``rule="dead-code:unused-import"``
- ``unused-property`` -> ``rule="dead-code:unused-property"``
- any other -> ``rule=f"dead-code:{type}"`` (still admissible)

The fingerprint is over (rule, path, symbol-name) where symbol-name is
the ``function`` / ``class`` / ``attr`` field, or the empty string when
vulture reports no name (a module-level finding). Empty names never
collapse: two unnamed findings in different files keep different
fingerprints because ``path`` is part of the input.

Liveness (HR-3): ``measured`` is true when vulture emitted at least one
finding entry. ``subjects`` is the count of files vulture scanned
(counted from the unique ``file`` values in the findings); when vulture
emits no findings at all, the script still emits ``measured: true`` and
``subjects: 0`` only when the input list is itself empty (a clean run),
otherwise it exits 2 (HR-3: a run that did not measure never says OK).

Exit codes: 0 success · 2 unreadable or malformed input.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

VALID_RULES = frozenset({
    "dead-code:unused-function",
    "dead-code:unused-variable",
    "dead-code:unused-class",
    "dead-code:unused-attribute",
    "dead-code:unused-import",
    "dead-code:unused-property",
})


class AdapterError(Exception):
    """A condition the adapter refuses to interpret; maps to exit 2."""


def parse_json(raw: bytes, label: str) -> object:
    try:
        return json.loads(raw.decode("utf-8"))
    except UnicodeDecodeError as exc:
        raise AdapterError(f"{label} is not valid UTF-8: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise AdapterError(f"{label} is not valid JSON: {exc}") from exc


def normalize_rule(vulture_type: str) -> str:
    rule = f"dead-code:{vulture_type}"
    return rule  # always accepted; the ratchet does not whitelist rules.


def symbol_name(entry: dict) -> str:
    """Pick the most specific name vulture gave us, else empty string."""
    for key in ("function", "class", "attr", "variable", "name"):
        value = entry.get(key)
        if isinstance(value, str) and value:
            return value
    return ""


def fingerprint_for(rule: str, path: str, symbol: str) -> str:
    """Stable content- and route-based identifier (HR-16)."""
    payload = f"{rule}|{path}|{symbol}".encode("utf-8")
    return hashlib.sha1(payload).hexdigest()


def normalize(raw: bytes) -> dict:
    """Convert vulture's JSON output into the ratchet's input shape."""
    data = parse_json(raw, "vulture output")
    if not isinstance(data, list):
        raise AdapterError(
            "vulture --format json output must be a JSON array of findings"
        )
    findings = []
    files_seen: set[str] = set()
    for i, entry in enumerate(data):
        if not isinstance(entry, dict):
            raise AdapterError(f"vulture finding {i} is not an object")
        path = entry.get("file")
        vulture_type = entry.get("type")
        if not isinstance(path, str) or not path:
            raise AdapterError(f"vulture finding {i}: 'file' must be a non-empty string")
        if not isinstance(vulture_type, str) or not vulture_type:
            raise AdapterError(f"vulture finding {i}: 'type' must be a non-empty string")
        rule = normalize_rule(vulture_type)
        symbol = symbol_name(entry)
        findings.append({
            "rule": rule,
            "path": path,
            "fingerprint": fingerprint_for(rule, path, symbol),
        })
        files_seen.add(path)
    return {
        "measured": True,
        "subjects": len(files_seen),
        "findings": findings,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="normalize vulture --format json output into the ratchet's input shape")
    parser.add_argument(
        "--vulture-output",
        default=None,
        help="path to the vulture JSON file (default: stdin)",
    )
    parser.add_argument(
        "--output", default=None,
        help="path to write the normalized JSON (default: stdout)",
    )
    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="backslashreplace")
    try:
        if args.vulture_output is None:
            if sys.stdin.isatty():
                raise AdapterError("no input: pipe vulture's JSON or pass --vulture-output")
            raw = sys.stdin.buffer.read()
        else:
            try:
                raw = Path(args.vulture_output).read_bytes()
            except OSError as exc:
                raise AdapterError(f"unreadable vulture output: {exc}") from exc
        normalized = normalize(raw)
    except AdapterError as exc:
        print(f"NORMALIZE ERROR: {exc}", file=sys.stderr)
        return 2
    rendered = json.dumps(normalized, indent=2)
    if args.output is None:
        print(rendered)
    else:
        try:
            Path(args.output).write_text(rendered + "\n", encoding="utf-8")
        except OSError as exc:
            print(f"NORMALIZE ERROR: cannot write output: {exc}", file=sys.stderr)
            return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())