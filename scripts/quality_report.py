#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — assets/scripts/quality_report.py
"""Aggregate every gate's indicator envelope into one report.

Gates answer pass/fail. Indicators answer "by how much, and which way is it moving" — which is
what you need to decide whether to spend a day on cleanup, and what a ratchet's target date is
measured against. This runs each gate in ``--json`` mode, merges the envelopes, writes
``quality-report.json``, and renders a markdown table into the CI step summary.

DETERMINISM: the report carries the commit SHA, never a wall-clock timestamp. Two runs over the
same commit must produce byte-identical output, otherwise "the report changed" stops meaning
"the code changed". Gates run in a fixed order and findings arrive pre-sorted from each gate.

Exit codes:
    0  every gate passed
    1  any gate failed, errored, or produced an unreadable envelope
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------

#: Fixed execution order (Hard Rule 13). Cheap and structural first, so the first failure a
#: developer sees is the one that changes the most downstream numbers.
GATES: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    ("layers", "check_layers.py", ()),
    ("complexity", "check_complexity.py", ()),
    ("mutation_sites", "check_mutation_sites.py", ()),
    ("dry", "check_dry.py", ()),
    # DA-13. The Makefile's `quality-report` help text has always claimed this gate
    # was aggregated here; it never was, so a security gate rejecting reintroduced
    # legacy crypto symbols ran nowhere in CI. Last on purpose: it is a symbol
    # walker, not a metric, so it neither consumes nor invalidates the numbers above.
    ("legacy_hashes", "check_legacy_hashes.py", ()),
    # L01 of expedientes-web-migration (#264). Same pattern as the legacy
    # crypto gate above: this is a structural walker, not a metric, so
    # it belongs after the numeric gates. It rejects reintroduction of Win32
    # network / OLE COM symbols (CAP-057) once the access backend is
    # retired. The other 6 retirement gates (CAP-058..063) land as
    # ``app/src/`` grows the corresponding legacy to retire.
    # The third tuple element passes extra args to the script: ``--gate``
    # selects which CAP's forbidden-symbol set to enforce; ``--gate-name``
    # sets the envelope ``gate`` key so the aggregator's ``failed_gates``
    # list reports the correct name even when the CAP id differs.
    (
        "legacy_retirement",
        "check_legacy_retirement.py",
        ("--gate", "CAP-057", "--gate-name", "legacy_retirement"),
    ),
)

#: The mutation gate is not here on purpose: it consumes a session database produced by a real
#: mutation run, which is far too slow for every PR. It runs on a schedule and reports through
#: the same envelope. Its absence from a PR report is expected; a PR report that claimed to
#: include it would be lying.

#: PR-scoped gates. They need a base ref and a PR body, so they only run with --include-pr-gates.
PR_GATES: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    ("branch_name", "check_branch_name.py", ()),
    ("pr_size", "check_pr_size.py", ()),
)

#: The published indicator set: what it means, and which direction is good. Everything here is
#: derived from a gate envelope — this table never computes anything itself.
INDICATOR_MEANINGS: dict[str, tuple[str, str]] = {
    "violations": ("architecture boundary violations", "lower"),
    "violation_classes": ("distinct violation kinds", "lower"),
    "max_complexity": ("highest cyclomatic complexity of any function", "lower"),
    "functions_over_ceiling": ("functions above the gate ceiling", "lower"),
    "functions_measured": ("functions the gate could measure", "higher"),
    "line_coverage_pct": ("statements executed by the suite", "higher"),
    "max_mutation_sites": ("largest mutation surface of any file", "lower"),
    "files_over_ceiling": ("files above the gate ceiling", "lower"),
    "total_mutation_sites": ("mutation surface of the whole package", "lower"),
    "mutation_score_pct": ("mutants the suite killed", "higher"),
    "survivors_total": ("mutants no test noticed", "lower"),
    "incompetent_ratio_pct": (
        "mutants that could not execute; high means a broken run",
        "lower",
    ),
    "mutants_measured": ("mutants the run actually evaluated", "higher"),
    "duplicate_groups": ("distinct duplicated blocks", "lower"),
    "duplicated_statements": ("statements inside a duplicated block", "lower"),
    "duplicated_ratio_pct": ("duplicated share of all statements", "lower"),
    "changed_lines": ("reviewable lines in this change", "lower"),
    "files_changed": ("files touched by this change", "lower"),
    "overridden": ("review budget explicitly overridden", "lower"),
    "conformant": ("branch name matches the convention", "higher"),
}

# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------


def _commit(root: Path) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), "rev-parse", "HEAD"],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    return result.stdout.strip() if result.returncode == 0 else "unknown"


def run_gate(scripts: Path, script: str, root: Path, extra: tuple[str, ...]) -> dict:
    result = subprocess.run(
        [sys.executable, str(scripts / script), "--root", str(root), "--json", *extra],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    try:
        envelope = json.loads(result.stdout)
    except json.JSONDecodeError:
        return {
            "gate": script.removesuffix(".py").removeprefix("check_"),
            "status": "error",
            "detail": (result.stderr or result.stdout or "no output").strip()[:500],
            "indicators": {},
            "ceilings": {},
            "findings": [],
        }
    envelope.setdefault("indicators", {})
    envelope.setdefault("ceilings", {})
    envelope.setdefault("findings", [])
    return envelope


def build_report(root: Path, envelopes: list[dict]) -> dict:
    indicators: dict[str, dict] = {}
    for envelope in envelopes:
        gate = envelope.get("gate", "unknown")
        ceilings = envelope.get("ceilings", {})
        for name, value in sorted(envelope.get("indicators", {}).items()):
            meaning, direction = INDICATOR_MEANINGS.get(name, (name, "lower"))
            indicators[f"{gate}.{name}"] = {
                "value": value,
                "ceiling": ceilings.get(name),
                "direction": direction,
                "meaning": meaning,
            }
    failed = [
        envelope.get("gate", "unknown")
        for envelope in envelopes
        if envelope.get("status") != "pass"
    ]
    return {
        "schema": "deterministic-quality-harness/quality-report/v1",
        # Commit, never a timestamp: the same commit must always produce the same report.
        "commit": _commit(root),
        "status": "pass" if not failed else "fail",
        "failed_gates": sorted(failed),
        "indicators": indicators,
        "gates": envelopes,
    }


def render_markdown(report: dict) -> str:
    icon = "PASS" if report["status"] == "pass" else "FAIL"
    lines = [
        f"## Quality indicators — {icon}",
        "",
        f"Commit `{report['commit'][:12]}`",
        "",
        "| Indicator | Value | Ceiling | Good direction | Meaning |",
        "|---|---:|---:|---|---|",
    ]
    for name, entry in report["indicators"].items():
        ceiling = "—" if entry["ceiling"] is None else entry["ceiling"]
        lines.append(
            f"| `{name}` | {entry['value']} | {ceiling} | {entry['direction']} | "
            f"{entry['meaning']} |"
        )

    failing = [envelope for envelope in report["gates"] if envelope.get("status") != "pass"]
    if failing:
        lines += ["", "### Failing gates", ""]
        for envelope in failing:
            gate = envelope.get("gate", "unknown")
            detail = envelope.get("detail")
            lines.append(
                f"**{gate}** — {envelope.get('status')}" + (f": {detail}" if detail else "")
            )
            for finding in envelope.get("findings", [])[:20]:
                lines.append(f"- `{finding['file']}:{finding['line']}` {finding['detail']}")
            lines.append("")
    return "\n".join(lines) + "\n"


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8.

    Python picks the output encoding from the platform locale, so the same gate emits different
    bytes on a Windows workstation (cp1252) and a Linux runner (utf-8) — and a non-encodable
    character crashes the write outright. A harness that claims determinism cannot let its own
    output depend on where it ran.
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="repository root")
    parser.add_argument(
        "--scripts",
        type=Path,
        default=None,
        help="directory holding the gate scripts (default: alongside this file)",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help="report path (default: <root>/quality-report.json)",
    )
    parser.add_argument(
        "--include-pr-gates",
        action="store_true",
        help="also run the PR-scoped gates (branch name, PR size)",
    )
    args = parser.parse_args(argv)

    root = args.root.resolve()
    scripts = (args.scripts or Path(__file__).resolve().parent).resolve()
    out = args.out or (root / "quality-report.json")

    selected = list(GATES) + (list(PR_GATES) if args.include_pr_gates else [])
    envelopes = [run_gate(scripts, script, root, extra) for _, script, extra in selected]
    report = build_report(root, envelopes)

    out.write_text(json.dumps(report, indent=2, sort_keys=False) + "\n", encoding="utf-8")
    markdown = render_markdown(report)
    print(markdown, end="")

    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as handle:
            handle.write(markdown)

    return 0 if report["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
