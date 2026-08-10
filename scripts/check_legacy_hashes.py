#!/usr/bin/env python3
# HARNESS-PROVENANCE: lanzadera-mvp DA-13 — scripts/check_legacy_hashes.py
"""DA-13 pin test harness — reject any legacy crypto symbols in `platform/src/`.

The launcher `TbUsuarios.Password` legacy stored SHA256-hex without salt
(`integrations-security.md` H1); reintroducing it through a re-merged legacy
helper would re-open the OWASP-2024 unacceptable surface. This walker rejects
the names the legacy era used so a silent reintroduction fails the suite.

The hardcoded symbol list is the contract. A symbol absent from the list is
NOT covered by this gate; a symbol present in the list is ALWAYS rejected
regardless of context. Phases 1+ keep the gate green by never introducing
these names. Phase 6 ships the full pytest integration
(`tests/lanzadera/auth/test_no_legacy_compat.py`); Phase 0 only wires the
script so the AST walker is importable and runs from the CLI.

Exit codes:
    0  no forbidden symbol found in any `.py` under `<root>/platform/src/`
    1  a forbidden symbol was found
"""

from __future__ import annotations

import argparse
import ast
import json
import sys
from pathlib import Path

# Symbols the legacy era used. Names are matched as Python identifiers
# (`ast.Name`, `ast.arg`, `ast.Attribute.attr`). String literals containing
# these names are NOT matched — comments, log strings and test fixtures are
# allowed to mention the names; only actual symbols are rejected.
FORBIDDEN_SYMBOLS: frozenset[str] = frozenset(
    {
        "legacy_hash",
        "verify_legacy",
        "sha256",
        "old_password",
        "migrate_password",
    }
)

ROOT_DIR = "app/src"

EXCLUDED_PARTS = frozenset({"__pycache__", ".venv", "venv", "build", "dist"})


def _iter_python_files(root: Path) -> list[Path]:
    source_root = root / ROOT_DIR
    if not source_root.is_dir():
        return []
    files: list[Path] = []
    for path in sorted(source_root.rglob("*.py")):
        if EXCLUDED_PARTS.intersection(path.parts):
            continue
        files.append(path)
    return files


def _collect_symbols(tree: ast.AST) -> list[tuple[str, int]]:
    """Collect identifiers and attribute names reachable from the tree.

    We deliberately walk both `ast.Name` (bare identifiers) and `ast.arg`
    (parameter names) and `ast.Attribute.attr` (member names). Function-level
    `def`, class-level `class`, and assignment targets are included so that a
    symbol used as `legacy_hash = ...` or `def legacy_hash(): ...` is also
    caught. Docstrings and string literals are skipped — this is a structural
    gate, not a content one.
    """
    symbols: list[tuple[str, int]] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Name):
            symbols.append((node.id, node.lineno))
        elif isinstance(node, ast.arg):
            symbols.append((node.arg, node.lineno))
        elif isinstance(node, ast.Attribute):
            symbols.append((node.attr, node.lineno))
        elif isinstance(node, ast.FunctionDef):
            symbols.append((node.name, node.lineno))
        elif isinstance(node, ast.AsyncFunctionDef):
            symbols.append((node.name, node.lineno))
        elif isinstance(node, ast.ClassDef):
            symbols.append((node.name, node.lineno))
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    symbols.append((target.id, node.lineno))
                elif isinstance(target, ast.Attribute):
                    symbols.append((target.attr, node.lineno))
    return symbols


def find_violations(root: Path) -> list[dict]:
    findings: list[dict] = []
    for path in _iter_python_files(root):
        display = str(path.relative_to(root)).replace("\\", "/")
        try:
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        except SyntaxError:
            continue
        for symbol, line in _collect_symbols(tree):
            if symbol in FORBIDDEN_SYMBOLS:
                findings.append(
                    {
                        "file": display,
                        "line": line,
                        "symbol": symbol,
                        "detail": f"forbidden symbol '{symbol}' (DA-13)",
                    }
                )
    return findings


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8 (Hard Rule 17)."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="repository root")
    parser.add_argument(
        "--json", action="store_true", help="emit the indicator envelope"
    )
    args = parser.parse_args(argv)

    root = args.root.resolve()
    violations = find_violations(root)

    if args.json:
        status = "pass" if not violations else "fail"
        print(
            json.dumps(
                {
                    "gate": "legacy_hashes",
                    "status": status,
                    "indicators": {"violations": len(violations)},
                    "ceilings": {"violations": 0},
                    "findings": violations,
                },
                indent=2,
            )
        )
        return 0 if status == "pass" else 1

    if violations:
        print(f"FAIL  {len(violations)} forbidden legacy symbol(s) found:")
        for entry in violations:
            print(f"        {entry['file']}:{entry['line']}  {entry['symbol']}")
        return 1

    print(f"OK    no legacy crypto symbols in {ROOT_DIR}/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
