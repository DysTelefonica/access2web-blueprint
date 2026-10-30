#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v2.2 — assets/scripts/check_test_classification.py
"""Test classification gate: every integration test file carries the integration marker.

pytest selection depends on the marker. A typical layout keeps unit tests in the default run and
excludes the integration directory from it (``addopts``), while a dedicated CI job selects the
directory with ``-m integration``. A file under the integration directory that carries no marker is
collected by neither run, so its tests silently never execute.

A file is clean when either a top-level ``pytestmark`` names the marker (a bare
``pytest.mark.<marker>`` or a list/tuple containing it) or every ``test_*`` function in the module
is decorated with it. A file with no test functions is not this gate's concern. There is no
baseline: an unmarked file is always a violation (Hard Rule 12 applies to ratchets; this rule has no
tolerated debt to shrink).

Hard Rule 18: a missing integration directory, or one with no ``test_*.py`` file, measured nothing
and fails closed instead of reporting a clean tree.

Domain-to-integration contracts (which modules must have an integration file, which atoms they must
contain) are consumer policy and do not belong in this asset.

Exit codes:
    0  every integration test file is marked
    1  at least one violation, or nothing was measured
"""

from __future__ import annotations

import argparse
import ast
import json
import sys
from pathlib import Path

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------

INTEGRATION_DIR = "tests/integration"
MARKER = "integration"

# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------


def _is_marker_expr(value: ast.expr, marker: str) -> bool:
    """True for the ``pytest.mark.<marker>`` attribute chain itself."""
    return (
        isinstance(value, ast.Attribute)
        and value.attr == marker
        and isinstance(value.value, ast.Attribute)
        and value.value.attr == "mark"
        and isinstance(value.value.value, ast.Name)
        and value.value.value.id == "pytest"
    )


def _has_module_marker(tree: ast.Module, marker: str) -> bool:
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if not any(isinstance(t, ast.Name) and t.id == "pytestmark" for t in node.targets):
            continue
        value = node.value
        if isinstance(value, (ast.List, ast.Tuple)):
            if any(_is_marker_expr(elt, marker) for elt in value.elts):
                return True
        elif _is_marker_expr(value, marker):
            return True
    return False


def file_violations(path: Path, marker: str) -> list[str]:
    """Return the violations of one integration test file (empty when clean)."""
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    except (OSError, UnicodeDecodeError, SyntaxError) as exc:
        return [f"{path.name}: file is not parseable: {exc}"]
    if _has_module_marker(tree, marker):
        return []
    tests = [
        node
        for node in tree.body
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        and node.name.startswith("test_")
    ]
    unmarked = sorted(
        node.name
        for node in tests
        if not any(_is_marker_expr(d, marker) for d in node.decorator_list)
    )
    if not unmarked:
        return []
    return [
        f"{path.name}: no pytest.mark.{marker} marker (neither a module-level pytestmark nor "
        f"a decorator on every test) -- unmarked: {', '.join(unmarked)}"
    ]


def check_tree(root: Path, integration_dir: str, marker: str) -> tuple[list[str], int]:
    """Return (violations, files_checked) for the integration directory under ``root``."""
    directory = root / integration_dir
    if not directory.is_dir():
        return [f"{integration_dir}: directory not found -- nothing was measured"], 0
    files = sorted(directory.glob("test_*.py"))
    if not files:
        return [f"{integration_dir}: no test_*.py file -- nothing was measured"], 0
    violations: list[str] = []
    for path in files:
        violations.extend(file_violations(path, marker))
    return violations, len(files)


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8 so the output does not depend on the platform locale."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", default=".", help="repository root (default: current directory)")
    parser.add_argument("--integration-dir", default=INTEGRATION_DIR)
    parser.add_argument("--marker", default=MARKER)
    parser.add_argument("--json", action="store_true", help="emit the indicator envelope")
    args = parser.parse_args(argv)

    violations, checked = check_tree(Path(args.root).resolve(), args.integration_dir, args.marker)

    if args.json:
        print(
            json.dumps(
                {
                    "gate": "test_classification",
                    "status": "fail" if violations else "pass",
                    "indicators": {"unmarked_files": len(violations), "files_checked": checked},
                    "ceilings": {"unmarked_files": 0},
                    "findings": [{"file": args.integration_dir, "line": 0, "detail": v} for v in violations],
                },
                indent=2,
            )
        )
        return 1 if violations else 0

    for violation in violations:
        print(f"FAIL  {violation}")
    if violations:
        print(f"check_test_classification: {len(violations)} violation(s).")
        return 1
    print(f"check_test_classification: OK ({checked} integration file(s) carry the marker)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
