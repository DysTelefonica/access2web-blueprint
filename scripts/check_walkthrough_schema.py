#!/usr/bin/env python3
"""HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — check_walkthrough_schema.py.

Walkthrough schema gate: a contract for `docs/03-aplicaciones/*/walkthrough-*.json`,
enforced.

The template lives at `skills/documentation-alan-style/references/templates/walkthrough.json.tmpl`
and documents the MUST / SHOULD / MAY classification per field. This gate enforces the
MUST fields. SHOULD and MAY are conventions, not enforcement.

Exit codes:
    0  all walkthroughs comply with the MUST contract
    1  one or more MUST fields are missing
    2  walkthrough JSON is malformed (parse error)

Advisory mode (default: ON): emit findings but exit 0. Once the 43 walkthroughs are
migrated to the v2.1 schema (#369 follow-up), flip the default by removing this
fallback or wire CI with `--strict`.
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

#: Top-level MUST fields every walkthrough JSON MUST contain.
MUST_TOP_FIELDS = (
    "group",
    "method_version",
    "forms_count",
    "forms",
    "summary",
)

#: Per-form MUST fields every form entry MUST contain.
MUST_FORM_FIELDS = (
    "formName",
    "sourcePath",
    "codegraph_summary",
    "ui",
    "geometry",
    "behavior",
    "verify_form_bindings",
    "unattended",
    "tool_warnings",
)

#: Per-form subfield MUST constraints. Maps parent field -> tuple of required sub-keys.
MUST_SUBFIELD_CONSTRAINTS = {
    "verify_form_bindings": ("status",),
}

#: Pattern matched by `git ls-files` over the repo to discover walkthrough files.
WALKTHROUGH_GLOB = "docs/03-aplicaciones/*/walkthrough-*.json"

# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------


def _git(*args: str) -> str:
    result = subprocess.run(
        ["git", *args], capture_output=True, text=True, check=False, encoding="utf-8"
    )
    if result.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout


def discover_walkthroughs(repo_root: Path) -> list[Path]:
    """Return the list of walkthrough JSON files tracked by git under docs/03-aplicaciones."""
    raw = _git("-C", str(repo_root), "ls-files", WALKTHROUGH_GLOB)
    files = []
    for line in raw.splitlines():
        line = line.strip()
        if line:
            files.append(repo_root / line)
    return sorted(files)


def validate_walkthrough(path: Path) -> tuple[list[dict], list[dict]]:
    """Validate a single walkthrough JSON.

    Returns a tuple (errors, warnings). Errors block the gate; warnings are advisory.
    Each issue is a dict with `file`, `line` (0 if unknown), `field`, and `detail`.
    """
    errors: list[dict] = []
    warnings: list[dict] = []

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        return (
            [
                {
                    "file": str(path.relative_to(path.parents[2])),
                    "line": exc.lineno,
                    "field": "<root>",
                    "detail": f"JSON parse error: {exc.msg}",
                }
            ],
            [],
        )

    for required in MUST_TOP_FIELDS:
        if required not in data:
            errors.append(
                {
                    "file": str(path.relative_to(path.parents[2])),
                    "line": 0,
                    "field": required,
                    "detail": f"MISSING top-level MUST field `{required}`",
                }
            )

    forms = data.get("forms")
    if isinstance(forms, list):
        for idx, form in enumerate(forms):
            if not isinstance(form, dict):
                errors.append(
                    {
                        "file": str(path.relative_to(path.parents[2])),
                        "line": 0,
                        "field": f"forms[{idx}]",
                        "detail": f"form entry #{idx} is not an object",
                    }
                )
                continue
            for required in MUST_FORM_FIELDS:
                if required not in form:
                    errors.append(
                        {
                            "file": str(path.relative_to(path.parents[2])),
                            "line": 0,
                            "field": f"forms[{idx}].{required}",
                            "detail": f"MISSING form MUST field `{required}`",
                        }
                    )
            for parent, subfields in MUST_SUBFIELD_CONSTRAINTS.items():
                child = form.get(parent)
                if isinstance(child, dict):
                    for sub in subfields:
                        if sub not in child:
                            errors.append(
                                {
                                    "file": str(path.relative_to(path.parents[2])),
                                    "line": 0,
                                    "field": f"forms[{idx}].{parent}.{sub}",
                                    "detail": f"MISSING `{parent}.{sub}` MUST subfield",
                                }
                            )

    return errors, warnings


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8 (see check_branch_name.py for rationale)."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--repo-root",
        default=os.environ.get("REPO_ROOT", "."),
        help="path to repo root (default: $REPO_ROOT or cwd)",
    )
    parser.add_argument("--json", action="store_true", help="emit the indicator envelope")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="exit non-zero on findings (default: advisory mode exits 0)",
    )
    args = parser.parse_args(argv)

    repo_root = Path(args.repo_root).resolve()
    try:
        walkthroughs = discover_walkthroughs(repo_root)
    except RuntimeError as exc:
        if args.json:
            print(
                json.dumps(
                    {
                        "gate": "walkthrough_schema",
                        "status": "error",
                        "detail": str(exc),
                    }
                )
            )
        else:
            print(f"FAIL  {exc}", file=sys.stderr)
        return 1

    all_errors: list[dict] = []
    files_with_errors = 0
    for path in walkthroughs:
        errors, _warnings = validate_walkthrough(path)
        if errors:
            files_with_errors += 1
            all_errors.extend(errors)

    passed = files_with_errors == 0
    advisory = not args.strict
    if args.json:
        envelope = {
            "gate": "walkthrough_schema",
            "status": "pass" if passed else ("fail" if not advisory else "advisory_fail"),
            "mode": "strict" if args.strict else "advisory",
            "indicators": {
                "walkthroughs_total": len(walkthroughs),
                "walkthroughs_with_errors": files_with_errors,
                "errors_total": len(all_errors),
            },
            "ceilings": {
                "walkthroughs_with_errors": 0,
                "errors_total": 0,
            },
            "findings": all_errors[:50] if not passed else [],
        }
        print(json.dumps(envelope, indent=2))
        return 0 if (passed or advisory) else 1

    if passed:
        print(f"OK    {len(walkthroughs)} walkthrough(s) comply with the MUST contract")
        return 0

    print(
        f"{'ADVISORY' if advisory else 'FAIL'}  {files_with_errors} of {len(walkthroughs)} "
        f"walkthrough(s) miss MUST fields ({len(all_errors)} total findings)"
    )
    for finding in all_errors[:20]:
        print(f"        {finding['file']} -> {finding['field']}: {finding['detail']}")
    if len(all_errors) > 20:
        print(f"        ... and {len(all_errors) - 20} more")
    if advisory:
        print("        (advisory mode: exiting 0; pass --strict to enforce)")
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
