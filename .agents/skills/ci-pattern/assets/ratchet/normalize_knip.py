#!/usr/bin/env python3
# ci-pattern asset — knip -> ratchet normalizer (DysTelefonica/team-skills#528)
"""Normalize `knip --reporter json` output into the ratchet's input shape.

The ratchet (`check_ratchet.py`) knows no measuring tool: it consumes a
list of findings with the stable identity ``(rule, path, fingerprint)``.
knip emits a JSON object with at least ``files`` and ``issues`` arrays;
each issue carries ``file``, ``type`` (one of ``files``, ``deps``,
``exports``, ``duplicates``, ``unresolved``, etc.) and a ``symbol``
field when the issue names one. This script maps each ``type`` to a
stable ``rule`` and computes ``fingerprint`` as the SHA-1 of
``rule|path|symbol-name`` (content- and route-based, never a mutable
field — HR-16).

Recognized knip ``type`` values:

- ``files``       -> ``rule="dead-code:unused-files"``      (entry-point files never imported)
- ``deps``        -> ``rule="dead-code:unused-deps"``       (declared but unused deps)
- ``devDeps``     -> ``rule="dead-code:unused-dev-deps"``
- ``exports``     -> ``rule="dead-code:unused-exports"``    (exports never imported)
- ``nsExports``   -> ``rule="dead-code:unused-ns-exports"``
- ``classMembers``-> ``rule="dead-code:unused-class-members"``
- ``duplicates``  -> ``rule="dead-code:duplicate-export"``
- ``unresolved``  -> ``rule="dead-code:unresolved-import"``
- ``binaries``    -> ``rule="dead-code:unused-binary"``
- any other      -> ``rule=f"dead-code:{type}"`` (admissible too).

The fingerprint is over (rule, path, symbol). When the issue has no
``symbol``, we fall back to the empty string: two unnamed findings in
different paths still differ on ``path``, and two unnamed findings in
the same path keep distinct identities only when ``symbol`` is
non-empty. That is an honest trade-off: knip's own report is
file-level for these cases, and re-running knip on a clean tree
yields the same shape.

Liveness (HR-3): ``measured`` is true and ``subjects`` counts the
unique files across ``files`` and ``issues``. When the input has no
``issues`` and no ``files``, ``measured: true`` and ``subjects: 0``
is the zero-debt target state; anything else with a malformed input
exits 2 (a run that did not measure never says OK).

Exit codes: 0 success · 2 unreadable or malformed input.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path


def normalize_rule(knip_type: str) -> str:
    canonical = {
        "files": "dead-code:unused-files",
        "deps": "dead-code:unused-deps",
        "devDeps": "dead-code:unused-dev-deps",
        "exports": "dead-code:unused-exports",
        "nsExports": "dead-code:unused-ns-exports",
        "classMembers": "dead-code:unused-class-members",
        "duplicates": "dead-code:duplicate-export",
        "unresolved": "dead-code:unresolved-import",
        "binaries": "dead-code:unused-binary",
    }
    return canonical.get(knip_type, f"dead-code:{knip_type}")


def parse_json(raw: bytes, label: str) -> object:
    try:
        return json.loads(raw.decode("utf-8"))
    except UnicodeDecodeError as exc:
        raise AdapterError(f"{label} is not valid UTF-8: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise AdapterError(f"{label} is not valid JSON: {exc}") from exc


class AdapterError(Exception):
    """A condition the adapter refuses to interpret; maps to exit 2."""


def symbol_name(entry: dict) -> str:
    for key in ("symbol", "name"):
        value = entry.get(key)
        if isinstance(value, str) and value:
            return value
    return ""


def fingerprint_for(rule: str, path: str, symbol: str) -> str:
    payload = f"{rule}|{path}|{symbol}".encode()
    return hashlib.sha1(payload).hexdigest()


def normalize(raw: bytes) -> dict:
    """Convert knip's JSON output into the ratchet's input shape."""
    data = parse_json(raw, "knip output")
    if not isinstance(data, dict):
        raise AdapterError("knip --reporter json output must be a JSON object")
    files = data.get("files")
    issues = data.get("issues")
    if not isinstance(files, list):
        raise AdapterError("knip output: 'files' must be a list")
    if not isinstance(issues, list):
        raise AdapterError("knip output: 'issues' must be a list")
    findings = []
    files_seen: set[str] = set()
    for path in files:
        if not isinstance(path, str) or not path:
            raise AdapterError("knip output: 'files' entries must be non-empty strings")
        files_seen.add(path)
    for i, entry in enumerate(issues):
        if not isinstance(entry, dict):
            raise AdapterError(f"knip issue {i} is not an object")
        path = entry.get("file")
        knip_type = entry.get("type")
        if not isinstance(path, str) or not path:
            raise AdapterError(f"knip issue {i}: 'file' must be a non-empty string")
        if not isinstance(knip_type, str) or not knip_type:
            raise AdapterError(f"knip issue {i}: 'type' must be a non-empty string")
        rule = normalize_rule(knip_type)
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


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(
        description="normalize knip --reporter json output into the ratchet's input shape")
    parser.add_argument(
        "--knip-output", default=None,
        help="path to the knip JSON file (default: stdin)",
    )
    parser.add_argument(
        "--output", default=None,
        help="path to write the normalized JSON (default: stdout)",
    )
    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="backslashreplace")
    try:
        if args.knip_output is None:
            if sys.stdin.isatty():
                raise AdapterError("no input: pipe knip's JSON or pass --knip-output")
            raw = sys.stdin.buffer.read()
        else:
            try:
                raw = Path(args.knip_output).read_bytes()
            except OSError as exc:
                raise AdapterError(f"unreadable knip output: {exc}") from exc
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
