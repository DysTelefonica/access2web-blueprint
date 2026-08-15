#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 — assets/scripts/check_mutation_sites.py
"""Static mutation-site density gate.

Counts AST-level mutation targets per file without running a single mutant. It is the cheap
proxy for "how much surface would a mutation run have to cover here", and it answers a question
the complexity ceiling cannot: a file can be full of small, simple functions and still be an
enormous mutation surface, because surface is a property of the file, not of any one function.

Upstream (`swarm-forge` `cleaner.prompt`) makes the consequence mandatory rather than advisory:
a changed file above the ceiling is SPLIT before handoff. The ceiling here is upstream's 100.
A mature codebase will need a ratchet to get there — that is what BASELINE is for.

Use ``--emit-baseline`` to generate the BASELINE block rather than hand-writing it; a ratchet you
have to type by hand is a ratchet nobody adopts.

Exit codes:
    0  no file above MAX_MUTATION_SITES outside BASELINE
    1  a new offender, a BASELINE entry that grew, or a BASELINE entry past its target date
"""

from __future__ import annotations

import argparse
import ast
import json
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------

SCAN_DIRS = ("app",)

#: Upstream's number. Above this, upstream splits the file before handoff rather than
#: negotiating. A legacy codebase ratchets down to it; it does not raise the ceiling to meet
#: the codebase.
MAX_MUTATION_SITES = 100

EXCLUDED_PARTS = frozenset({"__pycache__", ".venv", "venv", "build", "dist", "migrations"})


@dataclass(frozen=True)
class BaselineEntry:
    """A tolerated over-ceiling file with a mandatory exit plan (Hard Rule 12)."""

    sites: int
    target: int
    target_date: str  # ISO-8601, YYYY-MM-DD


# Measured on 2026-08-13 after issue #266 slice 1 split `coverage_gate.py` into
# the orchestrator file plus a helpers module. `coverage_gate.py` is now under
# the ceiling and no longer needs a BASELINE slot. `coverage_gate_helpers.py`
# carried the lookup + per-target resolution out of the gate and lands at 111
# sites — emitted by `--emit-baseline`, not hand-written: a ratchet you typed
# is a ratchet you got wrong. The ratchet target aligns with the DG-11 review
# date in `openspec/changes/architectural-guards-over-metrics/tasks.md`.
BASELINE: dict[str, BaselineEntry] = {
    "app/pytest_plugin/coverage_gate_helpers.py": BaselineEntry(
        sites=111, target=100, target_date="2027-02-13"
    ),
}

# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------

#: Nodes a mutation tool would rewrite. Kept deliberately close to what real mutation operators
#: target, so the static number tracks the real cost of a mutation run.
_MUTABLE_NODES = (
    ast.BinOp,
    ast.BoolOp,
    ast.UnaryOp,
    ast.Compare,
    ast.If,
    ast.IfExp,
    ast.While,
    ast.For,
    ast.Assert,
    ast.Raise,
    ast.Return,
    ast.AugAssign,
)


@dataclass(frozen=True)
class Measurement:
    file: str
    sites: int


def _count_sites(tree: ast.AST) -> int:
    total = 0
    for node in ast.walk(tree):
        if isinstance(node, _MUTABLE_NODES):
            total += 1
        elif isinstance(node, ast.Constant) and isinstance(node.value, (int, float, str, bool)):
            total += 1
        elif isinstance(node, ast.Call):
            total += len(node.args) + len(node.keywords)
    return total


def measure(root: Path) -> list[Measurement]:
    measurements: list[Measurement] = []
    for scan_dir in SCAN_DIRS:
        base = root / scan_dir
        if not base.is_dir():
            continue
        for path in sorted(base.rglob("*.py")):
            if EXCLUDED_PARTS.intersection(path.parts):
                continue
            try:
                tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            except SyntaxError:
                continue  # check_layers.py owns the unparseable-file failure
            display = str(path.relative_to(root)).replace("\\", "/")
            measurements.append(Measurement(file=display, sites=_count_sites(tree)))
    return sorted(measurements, key=lambda item: item.file)


def offenders_of(measurements: list[Measurement]) -> list[Measurement]:
    return [item for item in measurements if item.sites > MAX_MUTATION_SITES]


def evaluate(offenders: list[Measurement], today: date) -> tuple[int, list[str]]:
    lines: list[str] = []
    failed = False
    observed = {item.file: item for item in offenders}

    for key in sorted(observed):
        item = observed[key]
        allowance = BASELINE.get(key)
        if allowance is None:
            failed = True
            lines.append(
                f"FAIL  {item.file}: {item.sites} mutation sites, ceiling is {MAX_MUTATION_SITES}"
            )
            lines.append("        split the file before handoff, or ratchet it with a target date")
        elif item.sites > allowance.sites:
            failed = True
            lines.append(
                f"FAIL  {item.file}: grew to {item.sites} sites, BASELINE allows {allowance.sites}"
            )
        elif today.isoformat() > allowance.target_date and item.sites > allowance.target:
            failed = True
            lines.append(
                f"FAIL  {item.file}: BASELINE expired on {allowance.target_date} at "
                f"{item.sites} sites, target was {allowance.target}"
            )
        elif item.sites < allowance.sites:
            lines.append(
                f"NOTE  {item.file}: down to {item.sites} sites; lower the BASELINE to lock it in"
            )

    for key in sorted(BASELINE):
        if key not in observed:
            lines.append(f"NOTE  {key}: now under the ceiling; remove it from BASELINE")

    if not failed:
        lines.append(f"OK    every file at or below {MAX_MUTATION_SITES} mutation sites")
    return (1 if failed else 0), lines


def render_baseline(offenders: list[Measurement], today: date, horizon_days: int = 90) -> str:
    """Emit a BASELINE block. Every entry carries a target and a date; there is no other shape."""
    target_date = date.fromordinal(today.toordinal() + horizon_days).isoformat()
    lines = ["BASELINE: dict[str, BaselineEntry] = {"]
    for item in sorted(offenders, key=lambda entry: entry.file):
        lines.append(
            f'    "{item.file}": BaselineEntry(sites={item.sites}, '
            f'target={MAX_MUTATION_SITES}, target_date="{target_date}"),'
        )
    lines.append("}")
    return "\n".join(lines)


def build_report(measurements: list[Measurement], status: str) -> dict:
    offenders = offenders_of(measurements)
    return {
        "gate": "mutation_sites",
        "status": status,
        "indicators": {
            "max_mutation_sites": max((item.sites for item in measurements), default=0),
            "files_over_ceiling": len(offenders),
            "total_mutation_sites": sum(item.sites for item in measurements),
        },
        "ceilings": {"max_mutation_sites": MAX_MUTATION_SITES, "files_over_ceiling": 0},
        "findings": [
            {"file": item.file, "line": 0, "detail": f"{item.sites} mutation sites"}
            for item in sorted(offenders, key=lambda entry: entry.file)
        ],
    }


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8 so output bytes do not depend on the platform locale."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="repository root")
    parser.add_argument("--json", action="store_true", help="emit the indicator envelope")
    parser.add_argument(
        "--emit-baseline",
        action="store_true",
        help="print a BASELINE block for the current offenders instead of judging them",
    )
    args = parser.parse_args(argv)

    root = args.root.resolve()
    if not any((root / scan_dir).is_dir() for scan_dir in SCAN_DIRS):
        message = f"none of {SCAN_DIRS} found under {root}"
        if args.json:
            print(json.dumps({"gate": "mutation_sites", "status": "error", "detail": message}))
        else:
            print(f"FAIL  {message}", file=sys.stderr)
        return 1

    measurements = measure(root)

    if not measurements:
        # Hard Rule 18: a measurement that could not run must never score as a
        # perfect one. The subject set is empty, so every ceiling below is
        # trivially satisfied — the healthiest possible number for the least
        # healthy possible state. Guard the subject set, not just the path: the
        # missing-package case was already covered above; this is the one that
        # looks like success (harness v1.6).
        message = f"{list(SCAN_DIRS)} under {root} yielded no files to measure"
        if args.json:
            print(json.dumps({"gate": "mutation_sites", "status": "error", "detail": message}))
        else:
            print(f"FAIL  {message}", file=sys.stderr)
        return 1

    if args.emit_baseline:
        print(render_baseline(offenders_of(measurements), date.today()))
        return 0

    exit_code, lines = evaluate(offenders_of(measurements), date.today())
    if args.json:
        print(
            json.dumps(
                build_report(measurements, "pass" if exit_code == 0 else "fail"),
                indent=2,
            )
        )
    else:
        for line in lines:
            print(line)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
