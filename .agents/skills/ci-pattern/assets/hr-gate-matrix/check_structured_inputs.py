#!/usr/bin/env python3
# ci-pattern asset — structured inputs contract gate (HR-1, DysTelefonica/team-skills#245)
"""Contract gate for HR-1: gates read structured data, never scrape it.

Reads the HR→gate matrix, takes every ``gate`` entry's declared asset, and
scans its source for text-scraping anti-patterns of PR/issue government
data: ``gh pr view`` / ``gh issue view`` / ``gh pr list`` / ``gh issue
list`` without ``--json`` on the same line (host text output whose
title/body/labels would then be parsed as free text). Gates that consume
the structured payload (``--json``, event files, JSON documents) are clean;
pattern-validated subjects (conventional commits, branch names) are the
governed datum itself and stay allowed.

This file's own matrix entry is the check itself and is exempt from the
scan (its fixtures are test data). A gate asset missing on disk exits 2 —
HR-32's meta-gate owns existence; this gate refuses to judge what it
cannot read.

Exit codes: 0 clean · 1 findings · 2 invalid input or empty subject.
"""

from __future__ import annotations

import argparse
import ast
import json
import re
import sys
from pathlib import Path

class GateError(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


SCRAPING = re.compile(r"\bgh\s+(?:pr|issue)\s+(?:view|list)\b")
SCRAPING_PY = re.compile(r"['\"]gh['\"],\s*['\"](?:pr|issue)['\"],\s*['\"](?:view|list)['\"]")
SELF = Path(__file__).resolve()  # este propio gate también es gate declarado


def load_matrix(path: Path) -> list[dict]:
    try:
        doc = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise GateError(f"matrix {str(path)!r} unreadable: {exc}") from exc
    rules = doc.get("rules")
    if not isinstance(rules, list):
        raise GateError("matrix must carry a 'rules' list")
    return rules


def python_scrape_calls(text: str):
    """Líneas de llamadas python cuyo primer argumento es una lista/tupla
    literal que empieza por 'gh', 'pr'|'issue', 'view'|'list' sin '--json'
    en ningún elemento (HR-1: la evasión multilínea no pasa)."""
    import ast as _ast
    try:
        tree = _ast.parse(text)
    except SyntaxError:
        return []
    hits = []
    for node in _ast.walk(tree):
        if not isinstance(node, _ast.Call) or not node.args:
            continue
        first = node.args[0]
        if not isinstance(first, (_ast.List, _ast.Tuple)):
            continue
        parts = [e.value for e in first.elts if isinstance(e, _ast.Constant)
                 and isinstance(e.value, str)]
        if len(parts) >= 3 and parts[0] == "gh" and parts[1] in ("pr", "issue") \
                and parts[2] in ("view", "list") and "--json" not in parts:
            hits.append(node.lineno)
    return hits


def logical_lines(text: str):
    """Une continuaciones con backslash y cita el número de línea física inicial."""
    logical, start, pending = [], None, []
    for number, raw in enumerate(text.splitlines(), 1):
        if raw.endswith("\\") and not raw.endswith("\\\\"):
            pending.append((number, raw[:-1]))
            continue
        if pending:
            pending.append((number, raw))
            logical.append((pending[0][0], "".join(piece for _, piece in pending)))
            pending = []
        else:
            logical.append((number, raw))
    if pending:
        logical.append((pending[0][0], "".join(piece for _, piece in pending)))
    return logical


def main(argv):
    ap = argparse.ArgumentParser(description="Structured inputs contract gate (HR-1).")
    ap.add_argument("--skill-root", default=str(Path(__file__).resolve().parents[2]),
                    help="raíz de la skill (donde viven assets/ y references/)")
    ap.add_argument("--matrix", default=None, help="override de la ruta de la matriz")
    args = ap.parse_args(argv)
    skill_root = Path(args.skill_root).resolve()
    matrix_path = Path(args.matrix) if args.matrix else \
        skill_root / "references" / "hr-gate-matrix.json"
    try:
        rules = load_matrix(matrix_path)
    except GateError as exc:
        print(f"HR-1 ERROR: {exc}", file=sys.stderr)
        return 2
    try:
        own_rel = SELF.relative_to(skill_root).as_posix()
    except ValueError:
        own_rel = None  # fixture skill: este gate no es parte del sujeto
    gates = {}
    for rule in rules:
        if rule.get("enforcement") != "gate":
            continue
        asset = rule.get("asset")
        if not isinstance(asset, str) or not asset:
            continue
        if asset == own_rel:
            continue  # this check is its own subject; fixtures are test data
        gates.setdefault(asset, []).append(rule.get("id", "?"))
    if not gates:
        print("HR-1 ERROR: the matrix declares no gate assets to scan "
              "(empty subject)", file=sys.stderr)
        return 2
    report = []
    try:
        for asset in sorted(gates):
            path = skill_root / asset
            if not path.is_file():
                raise GateError(f"gate asset {asset!r} does not exist on disk")
            hrs = ", ".join(gates[asset])
            text = path.read_text(encoding="utf-8", errors="replace")
            for number, line in logical_lines(text):
                # HR-33/HR-8: la exención es un marcador explícito y auditable
                # en la línea, con motivo — nunca un silencio del escáner.
                if "hr-1: fixture" in line.lower():
                    continue
                if SCRAPING.search(line) and "--json" not in line:
                    report.append(f"HR-1 {asset}:{number}: scrapes host text output "
                                  f"({hrs}) — government data must come from "
                                  "the structured payload (--json / event file / git refs)")
            if path.suffix == ".py":
                for lineno in python_scrape_calls(text):
                    report.append(f"HR-1 {asset}:{lineno}: python call scrapes host text "
                                  f"output ({hrs}) — government data must come from the "
                                  "structured payload (--json / event file / git refs)")
    except GateError as exc:
        print(f"HR-1 ERROR: {exc}", file=sys.stderr)
        return 2
    if report:
        print(f"HR-1 FAIL: {len(report)} finding(s) — gates must read structured data")
        for line in report:
            print(f"  - {line}")
        return 1
    print(f"STRUCTURED-INPUTS OK: {len(gates)} gate asset(s) read structured "
          "data; no text scraping of government data.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
