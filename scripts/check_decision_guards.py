#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v1.6
# + architectural-guards-over-metrics DG-1, DG-2, DG-3, DG-4, DG-7, DG-13
# — assets/scripts/check_decision_guards.py
"""Decision-to-guard cross-check gate (slice 4, unwired).

Every architectural decision registered with an ``ID`` in a ``design.md``
table under ``openspec/`` MUST be covered by one of:

* a guard test under ``tests/**`` whose ``HARNESS-PROVENANCE`` header names
  the same ID (DG-4 — docstrings do not count),
* an entry in ``COVERED_BY`` pointing to a script already wired in
  ``quality_report.GATES`` (DG-5, with ``asserts={...}`` driving the constant
  drift check, DG-6),
* an entry in ``BASELINE`` whose ``target_date`` has not yet passed (DG-8).

Slice 4 ships the gate unwired; ``GATES`` does not name this script yet,
``COVERED_BY`` and ``BASELINE`` are empty, so every decision found in the
repository is reported as ``uncovered``. Slice 6 wires ``GATES`` and
populates the maps.

Exit codes:
    0  every decision is covered, escaped, or calendared
    1  any decision without coverage, or a malformed/absent input state
"""

from __future__ import annotations

import argparse
import ast
import json
import re
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

# --------------------------------------------------------------------------------------------
# CONFIGURATION — adjust this block when instantiating. Everything below it is mechanism.
# --------------------------------------------------------------------------------------------

OPENSPEC_ROOT = "openspec"
TESTS_ROOT = "tests"
SCRIPTS_ROOT = "scripts"

#: Prefixes that identify a decision ID. ``R-*`` is intentionally absent
#: (DG-2: the risk table does not carry decisions). ``D-*`` is absent because
#: the legacy decisions are written ``D8``, ``D88`` without the dash.
DECISION_ID_PREFIXES: tuple[str, ...] = ("DA-", "DG-", "QC-")

#: Path segments excluded from the walk. ``fixtures`` MUST stay here: the
#: two-tier fixtures deliberately plant fake ``HARNESS-PROVENANCE`` headers.
EXCLUDED_PARTS = frozenset(
    {
        "__pycache__",
        ".venv",
        "venv",
        "build",
        "dist",
        "migrations",
        "fixtures",
        "pytest_plugin",
    }
)

#: ``design.md`` files without a recognisable decision table, each with a
#: documented reason. Absent = fail-closed with ``no_decision_table`` (DG-3).
UNGOVERNED_DESIGNS: dict[str, str] = {}


@dataclass(frozen=True)
class Coverage:
    """Coverage declared by a structural gate already wired in ``GATES``."""

    gate: str  # script name relative to SCRIPTS_ROOT; MUST appear in GATES
    reason: str
    asserts: dict[str, object] = field(default_factory=dict)


#: Slice 4 ships empty — slice 6 populates it with DA-1, DA-13 and DG-11.
COVERED_BY: dict[str, Coverage] = {}


@dataclass(frozen=True)
class BaselineEntry:
    """A tolerated uncovered decision with a mandatory exit plan (DG-8)."""

    count: int  # 1 while the decision is uncovered
    target: int  # 0
    target_date: str  # ISO-8601, YYYY-MM-DD — mandatory


#: Slice 4 ships empty — slice 6 populates it for DA-2..DA-12.
BASELINE: dict[str, BaselineEntry] = {}

# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------

_DELIMITER_CELL = re.compile(r"^:?-{3,}:?$")
_DECISION_TOKEN = re.compile(r"^(DA|DG|QC)-\d+\b", re.IGNORECASE)
_DECISION_TOKEN_GLOBAL = re.compile(r"\b(?:DA|DG|QC)-\d+\b", re.IGNORECASE)
_HARNESS_MARKER = "HARNESS-PROVENANCE:"


def _normalise_cell(cell: str) -> str:
    """Strip whitespace, ``**`` and backticks; lowercase the result."""
    return cell.strip().strip("*").strip("`").strip().lower()


def _split_row(line: str) -> list[str]:
    """Split a single ``|`` row into cells, respecting ``\\|`` escapes and
    pipes inside backtick code spans (GFM behaviour)."""
    stripped = line.strip()
    if not stripped.startswith("|") or not stripped.endswith("|"):
        return []
    inner = stripped[1:-1]
    cells: list[str] = []
    current: list[str] = []
    in_code = False
    index = 0
    while index < len(inner):
        char = inner[index]
        if char == "\\" and index + 1 < len(inner) and inner[index + 1] == "|":
            current.append("|")
            index += 2
            continue
        if char == "`":
            in_code = not in_code
            current.append(char)
            index += 1
            continue
        if char == "|" and not in_code:
            cells.append("".join(current).strip())
            current = []
            index += 1
            continue
        current.append(char)
        index += 1
    cells.append("".join(current).strip())
    return cells


def _is_delimiter_row(cells: list[str]) -> bool:
    return bool(cells) and all(_DELIMITER_CELL.match(cell) for cell in cells)


def _is_decision_row(first_cell: str) -> bool:
    """DG-13: a decision row has exactly one ``(DA|DG|QC)-\\d+`` token in its
    first cell, anchored at the start of the cell."""
    normalised = _normalise_cell(first_cell)
    if not _DECISION_TOKEN.match(normalised):
        return False
    tokens = _DECISION_TOKEN_GLOBAL.findall(normalised)
    return len(tokens) == 1


def _header_is_id_column(header_cells: list[str]) -> bool:
    """DG-2: a table qualifies when its first header cell normalises to ``id``."""
    return bool(header_cells) and _normalise_cell(header_cells[0]) == "id"


def _extract_decision_token(first_cell: str) -> str | None:
    """Return the decision ID from a row's first cell, preserving source case."""
    match = _DECISION_TOKEN.match(first_cell.strip())
    return match.group(0) if match else None


def _iter_design_md(root: Path) -> list[Path]:
    """Every ``design.md`` under ``openspec/`` (archived included — DG-10)."""
    openspec = root / OPENSPEC_ROOT
    if not openspec.is_dir():
        return []
    return sorted(
        path for path in openspec.rglob("design.md") if not EXCLUDED_PARTS.intersection(path.parts)
    )


def _parse_design_file(path: Path) -> dict[str, tuple[str, int]]:
    """Parse a single ``design.md`` and return its decision rows."""
    decisions: dict[str, tuple[str, int]] = {}
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return decisions

    lines = text.splitlines()
    display = str(path).replace("\\", "/")

    index = 0
    while index < len(lines):
        cells = _split_row(lines[index])
        if len(cells) < 2 or not _header_is_id_column(cells):
            index += 1
            continue

        if index + 1 >= len(lines) or not _is_delimiter_row(_split_row(lines[index + 1])):
            index += 2
            continue

        cursor = index + 2
        column_count = len(cells)
        while cursor < len(lines):
            body_cells = _split_row(lines[cursor])
            if not body_cells:
                break
            if len(body_cells) != column_count:
                break
            first_cell = body_cells[0]
            if _is_decision_row(first_cell):
                decision_id = _extract_decision_token(first_cell)
                if decision_id is not None:
                    decisions.setdefault(
                        decision_id,
                        (display, cursor + 1),  # 1-indexed line
                    )
            cursor += 1

        index = cursor

    return decisions


def parse_decision_tables(root: Path) -> dict[str, tuple[str, int]]:
    """Walk every ``design.md`` under ``openspec/`` and return ``{id: (file, line)}``.

    Two ``design.md`` files contributing the same ID is permitted at this
    layer: the caller (``evaluate``) treats the collision as ``duplicate_id``
    (DG-7) using the raw multi-file shape.
    """
    root = root.resolve()
    collected: dict[str, tuple[str, int]] = {}
    for design_path in _iter_design_md(root):
        for decision_id, location in _parse_design_file(design_path).items():
            collected.setdefault(decision_id, location)
    return collected


def _collect_design_collisions(root: Path) -> dict[str, list[tuple[str, int]]]:
    """Same as ``parse_decision_tables`` but keeps every occurrence so
    ``evaluate`` can emit ``duplicate_id`` findings."""
    root = root.resolve()
    collisions: dict[str, list[tuple[str, int]]] = {}
    for design_path in _iter_design_md(root):
        for decision_id, location in _parse_design_file(design_path).items():
            collisions.setdefault(decision_id, []).append(location)
    return collisions


def _iter_python_files(root: Path) -> list[Path]:
    tests = root / TESTS_ROOT
    if not tests.is_dir():
        return []
    return sorted(
        path for path in tests.rglob("*.py") if not EXCLUDED_PARTS.intersection(path.parts)
    )


def _collect_provenance_block(lines: list[str]) -> list[str]:
    """Return the contiguous ``#`` lines starting with ``HARNESS-PROVENANCE:``."""
    block: list[str] = []
    in_block = False
    for line in lines:
        if not in_block:
            if line.lstrip().startswith("#") and _HARNESS_MARKER in line:
                in_block = True
                block.append(line)
            continue
        if line.lstrip().startswith("#"):
            block.append(line)
        else:
            break
    return block


def collect_guard_claims(root: Path) -> dict[str, list[str]]:
    """Walk ``tests/**`` and return ``{decision_id: [test_file, ...]}`` (DG-4)."""
    root = root.resolve()
    claims: dict[str, list[str]] = {}
    # Non-capturing outer group so ``findall`` returns the whole token.
    token_re = re.compile(r"(?:DA|DG|QC)-\d+\b")
    for path in _iter_python_files(root):
        try:
            text = path.read_text(encoding="utf-8")
        except OSError:
            continue
        block = _collect_provenance_block(text.splitlines())
        if not block:
            continue
        display = str(path.relative_to(root)).replace("\\", "/")
        for token in token_re.findall("\n".join(block)):
            claims.setdefault(token, []).append(display)
    return claims


def _read_quality_report_gates(root: Path) -> set[str]:
    """Parse the GATES tuple out of ``scripts/quality_report.py`` by AST."""
    script = root / SCRIPTS_ROOT / "quality_report.py"
    if not script.is_file():
        return set()
    try:
        tree = ast.parse(script.read_text(encoding="utf-8"), filename=str(script))
    except SyntaxError:
        return set()
    gates: set[str] = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign):
            continue
        if not any(getattr(target, "id", None) == "GATES" for target in node.targets):
            continue
        value = node.value
        if not isinstance(value, ast.Tuple):
            continue
        for element in value.elts:
            if isinstance(element, ast.Tuple) and element.elts:
                first = element.elts[0]
                if isinstance(first, ast.Constant) and isinstance(first.value, str):
                    gates.add(first.value)
    return gates


def _read_constant(root: Path, script_name: str, name: str) -> object | None:
    """Read a top-level module constant by AST. Returns ``None`` if absent."""
    script_path = root / SCRIPTS_ROOT / script_name
    if not script_path.is_file():
        return None
    try:
        tree = ast.parse(script_path.read_text(encoding="utf-8"), filename=str(script_path))
    except SyntaxError:
        return None
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if not any(getattr(target, "id", None) == name for target in node.targets):
            continue
        value = node.value
        if isinstance(value, ast.Constant):
            return value.value
        if isinstance(value, ast.List):
            return [elt.value for elt in value.elts if isinstance(elt, ast.Constant)]
        if isinstance(value, ast.Tuple):
            return tuple(elt.value for elt in value.elts if isinstance(elt, ast.Constant))
        if isinstance(value, ast.Dict):
            return {
                k.value: v.value
                for k, v in zip(value.keys, value.values, strict=True)
                if isinstance(k, ast.Constant) and isinstance(v, ast.Constant)
            }
        if isinstance(value, ast.Set):
            return frozenset(elt.value for elt in value.elts if isinstance(elt, ast.Constant))
    return None


def validate_coverage(
    root: Path,
    decisions: dict[str, tuple[str, int]],
    claims: dict[str, list[str]],
    covered_by: dict[str, Coverage],
    baseline: dict[str, BaselineEntry],
    gates: set[str] | None = None,
    today: date | None = None,
) -> tuple[list[dict], list[dict], list[dict], list[dict]]:
    """Cross-check inputs and return ``(findings, notes, uncovered, expired)``."""
    if gates is None:
        gates = _read_quality_report_gates(root)
    if today is None:
        today = date.today()

    findings: list[dict] = []
    notes: list[dict] = []
    uncovered: list[dict] = []
    expired: list[dict] = []

    # DG-7: global uniqueness across the tree.
    collisions = _collect_design_collisions(root)
    for decision_id, occurrences in collisions.items():
        if len(occurrences) > 1:
            file_lines = ", ".join(f"{file}:{line}" for file, line in occurrences)
            findings.append(
                {
                    "key": "duplicate_id",
                    "id": decision_id,
                    "file": occurrences[0][0],
                    "line": occurrences[0][1],
                    "detail": (
                        f"decision {decision_id} appears in {len(occurrences)} files: {file_lines}"
                    ),
                }
            )

    # DG-3: empty subject set fails closed (Hard Rule 18).
    if not decisions:
        findings.append(
            {
                "key": "no_decision_table",
                "id": None,
                "file": str(root),
                "line": 0,
                "detail": "no decision table found under openspec/; gate fails closed",
            }
        )

    # DG-10: a guard header citing an unknown ID fails. Independent of the
    # empty-subject-set check.
    all_known = set(decisions) | set(covered_by) | set(baseline)
    for token, files in claims.items():
        if token not in all_known:
            findings.append(
                {
                    "key": "orphan_guard",
                    "id": token,
                    "file": files[0],
                    "line": 0,
                    "detail": (
                        f"HARNESS-PROVENANCE cites {token} but no decision or "
                        "BASELINE entry exists for it"
                    ),
                }
            )

    if not decisions:
        return findings, notes, uncovered, expired

    for decision_id, (file_path, line_no) in sorted(decisions.items()):
        if decision_id in claims:
            continue
        if decision_id in covered_by:
            coverage = covered_by[decision_id]
            if coverage.gate not in gates:
                findings.append(
                    {
                        "key": "stale_coverage",
                        "id": decision_id,
                        "file": file_path,
                        "line": line_no,
                        "detail": (
                            f"COVERED_BY names '{coverage.gate}' which is absent from "
                            "quality_report.GATES"
                        ),
                    }
                )
                uncovered.append({"id": decision_id, "file": file_path, "line": line_no})
                continue
            drift = False
            for constant_name, expected_value in coverage.asserts.items():
                actual = _read_constant(root, coverage.gate, constant_name)
                if actual != expected_value:
                    findings.append(
                        {
                            "key": "constant_drift",
                            "id": decision_id,
                            "file": file_path,
                            "line": line_no,
                            "detail": (
                                f"{coverage.gate}.{constant_name} expected "
                                f"{expected_value!r} but is {actual!r}"
                            ),
                        }
                    )
                    drift = True
                    break
            if drift:
                uncovered.append({"id": decision_id, "file": file_path, "line": line_no})
                continue
            continue
        if decision_id in baseline:
            entry = baseline[decision_id]
            if today.isoformat() > entry.target_date and entry.count > entry.target:
                findings.append(
                    {
                        "key": "baseline_expired",
                        "id": decision_id,
                        "file": file_path,
                        "line": line_no,
                        "detail": (
                            f"BASELINE expired on {entry.target_date} with count={entry.count}, "
                            f"target={entry.target}"
                        ),
                    }
                )
                expired.append({"id": decision_id, "file": file_path, "line": line_no})
                continue
            notes.append(
                {
                    "key": "baseline_open",
                    "id": decision_id,
                    "file": file_path,
                    "line": line_no,
                    "detail": (f"BASELINE active until {entry.target_date}, count={entry.count}"),
                }
            )
            continue
        findings.append(
            {
                "key": "uncovered",
                "id": decision_id,
                "file": file_path,
                "line": line_no,
                "detail": (
                    f"decision {decision_id} has no guard, no COVERED_BY, and no BASELINE entry"
                ),
            }
        )
        uncovered.append({"id": decision_id, "file": file_path, "line": line_no})

    for decision_id in sorted(baseline):
        if decision_id not in decisions:
            notes.append(
                {
                    "key": "baseline_id_unknown",
                    "id": decision_id,
                    "file": str(root / OPENSPEC_ROOT),
                    "line": 0,
                    "detail": (
                        f"BASELINE entry for {decision_id} has no matching decision table row"
                    ),
                }
            )

    return findings, notes, uncovered, expired


def evaluate(
    decisions: dict[str, tuple[str, int]],
    findings: list[dict],
) -> tuple[int, list[str]]:
    """Render the human-readable report from the verdict lists."""
    lines: list[str] = []
    failed = bool(findings)

    for finding in findings:
        lines.append(
            f"FAIL  [{finding['key']}] {finding.get('id') or '—'}: "
            f"{finding['detail']}  ({finding['file']}:{finding['line']})"
        )

    if not failed:
        lines.append("OK    every decision is covered, escaped, or calendared")
    return (1 if failed else 0), lines


def build_report(
    decisions: dict[str, tuple[str, int]],
    claims: dict[str, list[str]],
    findings: list[dict],
    notes: list[dict],
    status: str,
    designs_scanned: int,
) -> dict:
    """Produce the envelope JSON consumed by ``quality_report.py``."""
    uncovered_ids = {f.get("id") for f in findings if f.get("key") == "uncovered"}
    orphan_ids = {f.get("id") for f in findings if f.get("key") == "orphan_guard"}
    return {
        "gate": "decision_guards",
        "status": status,
        "indicators": {
            "decisions_total": len(decisions),
            "decisions_uncovered": len([i for i in uncovered_ids if i]),
            "guards_total": len(claims),
            "orphan_guards": len([i for i in orphan_ids if i]),
            "designs_scanned": designs_scanned,
            "designs_ungoverned": 0,
        },
        "ceilings": {
            "decisions_uncovered": 0,
            "orphan_guards": 0,
            "designs_ungoverned": 0,
        },
        "findings": [
            {"file": f["file"], "line": f["line"], "detail": f["detail"]} for f in findings
        ]
        + [{"file": n["file"], "line": n["line"], "detail": n["detail"]} for n in notes],
    }


def _pin_output_encoding() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="repository root")
    parser.add_argument("--json", action="store_true", help="emit the indicator envelope")
    args = parser.parse_args(argv)

    root = args.root.resolve()
    decisions = parse_decision_tables(root)
    claims = collect_guard_claims(root)
    findings, notes, _uncovered, _expired = validate_coverage(
        root, decisions, claims, COVERED_BY, BASELINE
    )
    designs_scanned = len(_iter_design_md(root))
    exit_code, lines = evaluate(decisions, findings)

    if args.json:
        envelope = build_report(
            decisions,
            claims,
            findings,
            notes,
            "pass" if exit_code == 0 else "fail",
            designs_scanned,
        )
        print(json.dumps(envelope, indent=2))
    else:
        for line in lines:
            print(line)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
