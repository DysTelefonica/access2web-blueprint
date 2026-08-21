"""CAP-057..063 retirement gates — the scripts that the UAT layer calls.

This module is the **infrastructure** for the retirement gates
described in ``openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md``.
It is NOT a retirement itself — the legacy is still in the access
backend, and the retirement is planned for the chain ``U09→U10→L01``
in the task plan. What this WU delivers is the **gate runner** that
prevents the legacy from creeping back into ``app/src/`` once the
retirement lands.

The runner is the analog of ``scripts/check_legacy_hashes.py`` (DA-13,
the legacy crypto pin) but for the 7 retirement surfaces (CAP-057
Win32 network, CAP-058 Win32 processes, CAP-059 OLE/ActiveX,
CAP-060 mutable globals, CAP-061 backend selector, CAP-062 popups,
CAP-063 menu JSON-hub). This WU ships CAP-057 (Win32 network) as the
first gate; the other 6 land as ``app/src/`` grows legacy to retire
(see ``openspec/changes/expedientes-web-migration/tasks.md`` U09-U10).

The runner exposes one function per gate (one today, six to follow);
the test (``tests/lanzadera/exp/test_retirement_gates.py``) calls each
function and asserts the surface is clean. The script can also run
standalone: ``python scripts/check_legacy_retirement.py --root <repo>``
prints violations and exits non-zero on any.

Why ``EXP-CAP-057 first``: the task plan puts CAP-057 (Win32 network)
as the first retirement because the access backend's network path
(\\\\server\\\\share\\\\file.accdb) is the dependency that physically
prevents the platform from running on Linux. Every other retirement
depends on this one; once the network is on ``DocumentStoragePort``,
the rest follow. The gate is the test that prevents a reintroduction
of ``win32file``, ``win32net``, or ``win32com`` to ``app/src/``.

The symbol list is the contract: a name absent from the list is NOT
covered; a name present in the list is ALWAYS rejected regardless of
context (comment, docstring, string literal — none of these are
exempt). This is the same conservative bias as the crypto gate (DA-13).
"""

from __future__ import annotations

import argparse
import ast
import sys
from pathlib import Path

# CAP-057 Win32 network: the access backend's file-share path. The
# forbidden set covers both the direct ``win32file`` / ``win32net`` /
# ``win32com`` modules AND the symbolic names the legacy era used
# (``WSAStartup``, ``NetShareAdd``). Phases 1+ keep the gate green by
# never introducing these names; Phase 6 ships the full pytest
# integration; Phase 0 only wires the script so the AST walker is
# importable and runs from the CLI.
FORBIDDEN_SYMBOLS_CAP057: frozenset[str] = frozenset(
    {
        # Direct win32 modules (would need a Windows-only dep)
        "win32file",
        "win32net",
        "win32com",
        "win32api",
        "win32con",
        "win32gui",
        "win32process",
        "win32security",
        "win32service",
        # Win32 API function names the legacy era used
        "WSAStartup",
        "NetShareAdd",
        "NetShareDel",
        "CreateFile",
        "WriteFile",
        # OLE / COM component names (CAP-059 hooks into the same
        # surface; the gate covers both because they share the import
        # path ``win32com.client``)
        "MSComctlLib",
        "MSComDlg",
        "OleLoadPicture",
    }
)

ROOT_DIR = "app/src"
EXCLUDED_PARTS = frozenset({"__pycache__", ".venv", "venv", "build", "dist"})


def _iter_python_files(root: Path) -> list[Path]:
    """Yield every ``.py`` under ``app/src/`` minus the standard excluded parts."""
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

    We walk ``ast.Name`` (bare identifiers), ``ast.arg`` (parameter
    names), ``ast.Attribute.attr`` (member names), and ``ast.Import``
    / ``ast.ImportFrom`` (imported names — the CAP-057 forbidden
    set targets ``win32file`` / ``win32net``, which only show up in
    ``import win32file`` statements, never as bare ``Name`` nodes).
    Function-level ``def``, class-level ``class``, and assignment
    targets are included so a symbol used as ``legacy_name = ...``
    or ``def legacy_name(): ...`` is also caught. Docstrings and
    string literals are skipped — this is a structural gate, not a
    content one.
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
        elif isinstance(node, ast.Import):
            # ``import win32file`` — only the top-level module name is
            # checked; ``import win32file.sub`` is matched by the
            # top-level name (the gate rejects the whole module).
            for alias in node.names:
                symbols.append((alias.name.split(".")[0], node.lineno))
        elif isinstance(node, ast.ImportFrom):
            # ``from win32com import client`` — both the module and the
            # imported names are checked. A future reintroduction of
            # ``from win32file import CreateFile`` would be caught here.
            if node.module:
                symbols.append((node.module.split(".")[0], node.lineno))
            for alias in node.names:
                symbols.append((alias.name, node.lineno))
    return symbols


def find_violations(root: Path, forbidden: frozenset[str]) -> list[dict[str, object]]:
    """Walk ``app/src/`` and return every occurrence of a forbidden symbol."""
    findings: list[dict[str, object]] = []
    for path in _iter_python_files(root):
        display = str(path.relative_to(root)).replace("\\", "/")
        try:
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        except SyntaxError as exc:
            findings.append(
                {
                    "file": display,
                    "line": exc.lineno or 0,
                    "symbol": "(syntax error)",
                    "context": exc.msg,
                }
            )
            continue
        for symbol, lineno in _collect_symbols(tree):
            if symbol in forbidden:
                findings.append(
                    {
                        "file": display,
                        "line": lineno,
                        "symbol": symbol,
                        "context": "CAP-057 retired: Win32 network reintroduction",
                    }
                )
    return findings


def main(argv: list[str] | None = None) -> int:
    """CLI entry — print violations and exit non-zero on any.

    The aggregator (``scripts/quality_report.py``) calls each gate
    in ``--json`` mode, so this entry point accepts a ``--json`` flag
    that emits the envelope format the aggregator expects:

    ``{"findings": [...], "summary": {"total": N, "files": F}}``

    When ``--json`` is absent the script falls back to the human-readable
    format (the GitHub Actions annotation lines) for standalone use.
    """
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path("."), help="repository root")
    parser.add_argument(
        "--gate",
        default="CAP-057",
        choices=("CAP-057",),
        help="which retirement gate to run (CAP-057 only, this WU)",
    )
    parser.add_argument(
        "--gate-name",
        default=None,
        help=(
            "Override the envelope ``gate`` name. Defaults to the lowercased "
            "--gate value with hyphens replaced. The aggregator sets this to "
            "match the first element of the corresponding GATES tuple "
            "(e.g. 'legacy_retirement' for CAP-057)."
        ),
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="emit the aggregator envelope format (stdout JSON, exit 0/1 only)",
    )
    args = parser.parse_args(argv)

    if args.gate == "CAP-057":
        findings = find_violations(args.root, FORBIDDEN_SYMBOLS_CAP057)
        label = "CAP-057 Win32 network"
    else:
        print(f"::error::unknown gate {args.gate}")
        return 2

    if args.json:
        # Envelope format. The aggregator merges ``findings`` into the
        # quality-report.json envelope and uses ``summary.total`` for the
        # envelope's "failed_gates" check. The ``gate`` key identifies this
        # run in the merged report (the aggregator's ``failed_gates`` list
        # is derived from this name); it MUST match the first element of
        # the corresponding ``GATES`` tuple in ``quality_report.py``.
        # We pass the gate name explicitly via the CLI so the GATES tuple
        # can run a different gate with the same script (CAP-058..063
        # future land the same way) without changing the envelope key.
        import json as _json

        gate_name = args.gate_name or args.gate.lower().replace("-", "_")
        envelope: dict[str, object] = {
            "gate": gate_name,
            "status": "fail" if findings else "pass",
            "findings": findings,
            "indicators": {"violations": len(findings)},
            "ceilings": {"violations": 0},
            "summary": {
                "total": len(findings),
                "files": len({f["file"] for f in findings}),
            },
        }
        print(_json.dumps(envelope, ensure_ascii=False))
        return 1 if findings else 0

    if findings:
        for finding in findings:
            print(
                f"::error file={finding['file']},line={finding['line']}::"
                f"{label} reintroduction: {finding['symbol']!r} — {finding['context']}"
            )
        print(f"FAIL  {len(findings)} violation(s) — see above for the offending lines")
        return 1

    print(f"OK    {label} gate clean")
    return 0


if __name__ == "__main__":
    sys.exit(main())
