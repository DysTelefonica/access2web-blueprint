#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — scripts/check_test_classification.py
"""Test classification gate (--dry-run default).

Enforces three rules of the taxonomy in ``docs/testing/testing-strategy.md``:

- **HR-2**: ``application/`` tests must not import ``unittest.mock.MagicMock``.
- **HR-6**: ``delivery/`` tests touching ``admin_routes`` must declare
  ``auth_bypass`` fixture or ``@pytest.mark.requires_auth`` marker.
- **HR-8**: test files at ``tests/lanzadera/test_*.py`` (root, not under a
  layer subdir) must be in ``WIRE_ALLOWLIST``.

Default mode is ``--dry-run``: print violations to stderr, exit 0.
Use ``--strict`` to exit 1 on violations. See ``docs/testing/testing-strategy.md``
§Pieza 4 for the rollout plan (one sprint of dry-run, then ``--strict`` in CI).
"""

from __future__ import annotations

import argparse
import ast
import re
import sys
from dataclasses import dataclass
from pathlib import Path

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------

DEFAULT_TESTS_ROOT = Path("tests/lanzadera")

#: Allowlist of files at tests/lanzadera/test_*.py (root). These are wiring/meta
#: tests (CI workflow pins, gate wiring, smoke tests). Add a name ONLY when the
#: test exercises the test infrastructure itself; product tests belong under
#: tests/lanzadera/<layer>/test_*.py.
WIRE_ALLOWLIST: frozenset[str] = frozenset(
    {
        "test_assume_in_office.py",
        "test_bootstrap_adapter.py",
        "test_branch_name_wiring.py",
        "test_complexity_wiring.py",
        "test_coverage_gate_class_methods_e2e.py",
        "test_coverage_gate_plugin.py",
        "test_crap_retired_guard.py",
        "test_decision_guards_wiring.py",
        "test_dry_wiring.py",
        "test_env_secret_manager_adapter.py",
        "test_layers_wiring.py",
        "test_pr_size_wiring.py",
        "test_quality_report_smoke.py",
        "test_secret_manager_contract.py",
        "test_ttl_cache_adapter.py",
    }
)

#: Patterns that count as "touches admin_routes" for HR-6.
ADMIN_ROUTES_PATTERNS: tuple[re.Pattern[str], ...] = (
    re.compile(r"\badmin_routes\b"),
    re.compile(r"\bfrom\s+app\.src\.modules\.lanzadera\.delivery\.http\.admin_routes\b"),
)

# --------------------------------------------------------------------------------------------
# Violation model
# --------------------------------------------------------------------------------------------


@dataclass(frozen=True)
class Violation:
    """A single violation of the test classification taxonomy."""

    rule: str
    file: str
    detail: str

    def __str__(self) -> str:
        return f"{self.rule}  {self.file}  {self.detail}"


# --------------------------------------------------------------------------------------------
# RULE 1 — HR-2: application/ tests must not import unittest.mock.MagicMock
# --------------------------------------------------------------------------------------------


def rule_no_magicmock_in_application(tests_root: Path) -> list[Violation]:
    """Detect ``unittest.mock.MagicMock`` imports under ``application/``.

    The contract: use case tests substitute dependencies via ``FakeFixtures``.
    ``MagicMock`` loses the signal of the port's contract and silently passes
    when the contract changes. See skill ``lanzadera-testing-strategy`` HR-2.
    """
    out: list[Violation] = []
    application_dir = tests_root / "application"
    if not application_dir.is_dir():
        return out
    for path in sorted(application_dir.glob("test_*.py")):
        try:
            source = path.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError) as exc:
            out.append(
                Violation(
                    rule="R1",
                    file=str(path),
                    detail=f"cannot read file: {exc}",
                )
            )
            continue
        if "MagicMock" not in source:
            continue
        # Confirm via AST that MagicMock is imported (not just mentioned in a comment).
        if _has_magicmock_import(source):
            out.append(
                Violation(
                    rule="R1",
                    file=str(path),
                    detail=(
                        "MagicMock is forbidden in application/ tests. "
                        "Use FakeFixtures from tests/lanzadera/_fakes.py. See skill HR-2."
                    ),
                )
            )
    return out


def _has_magicmock_import(source: str) -> bool:
    try:
        tree = ast.parse(source)
    except SyntaxError:
        return False
    for node in ast.walk(tree):
        if not isinstance(node, ast.ImportFrom):
            continue
        if node.module and node.module.startswith("unittest.mock"):
            for alias in node.names:
                if alias.name == "MagicMock":
                    return True
        if node.module == "mock" and any(alias.name == "MagicMock" for alias in node.names):
            return True
    return False


# --------------------------------------------------------------------------------------------
# RULE 2 — HR-6: delivery/ tests touching admin_routes declare auth_bypass or requires_auth
# --------------------------------------------------------------------------------------------


def rule_auth_bypass_in_delivery(tests_root: Path) -> list[Violation]:
    """Detect delivery tests that touch ``admin_routes`` without declaring an auth gate.

    Two acceptable gates:
    - A fixture named ``auth_bypass`` declared in the same file.
    - A ``@pytest.mark.requires_auth`` marker on the test function.

    Without one of these, the test runs without auth and may pass while the
    production gate (M02 / A01-A03) would block it. See skill HR-6.
    """
    out: list[Violation] = []
    delivery_dir = tests_root / "delivery"
    if not delivery_dir.is_dir():
        return out
    for path in sorted(delivery_dir.glob("test_*.py")):
        try:
            source = path.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError) as exc:
            out.append(
                Violation(
                    rule="R2",
                    file=str(path),
                    detail=f"cannot read file: {exc}",
                )
            )
            continue
        if not any(pattern.search(source) for pattern in ADMIN_ROUTES_PATTERNS):
            continue
        # File touches admin_routes — does it declare the auth gate?
        if "auth_bypass" in source or "@pytest.mark.requires_auth" in source:
            continue
        out.append(
            Violation(
                rule="R2",
                file=str(path),
                detail=(
                    "delivery/ test imports admin_routes but does not declare "
                    "auth_bypass fixture or @pytest.mark.requires_auth marker. "
                    "See skill HR-6."
                ),
            )
        )
    return out


# --------------------------------------------------------------------------------------------
# RULE 3 — HR-8: root-level tests must be in the wire allowlist
# --------------------------------------------------------------------------------------------


def rule_root_path_layer(tests_root: Path) -> list[Violation]:
    """Detect root-level test files that are not in the wire allowlist.

    Every product test must live under ``tests/lanzadera/<layer>/test_*.py``.
    A root-level file is a layer-violation unless it is one of the WIRE_ALLOWLIST
    entries (CI wiring, gate tests). See skill HR-8.
    """
    out: list[Violation] = []
    if not tests_root.is_dir():
        return out
    for path in sorted(tests_root.glob("test_*.py")):
        if path.name in WIRE_ALLOWLIST:
            continue
        out.append(
            Violation(
                rule="R3",
                file=str(path),
                detail=(
                    "test file at tests/lanzadera/test_*.py (root) is not in "
                    "WIRE_ALLOWLIST. Move to tests/lanzadera/<layer>/test_*.py "
                    "or add to WIRE_ALLOWLIST in scripts/check_test_classification.py. "
                    "See skill HR-8."
                ),
            )
        )
    return out


# --------------------------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------------------------


def _pin_output_encoding() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        default=str(DEFAULT_TESTS_ROOT),
        help=f"path to the tests directory (default: {DEFAULT_TESTS_ROOT})",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="exit 1 on violations (default is --dry-run: print only, exit 0)",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="emit the indicator envelope as JSON",
    )
    args = parser.parse_args(argv)

    tests_root = Path(args.root)
    violations: list[Violation] = []
    violations.extend(rule_no_magicmock_in_application(tests_root))
    violations.extend(rule_auth_bypass_in_delivery(tests_root))
    violations.extend(rule_root_path_layer(tests_root))

    if args.json:
        import json

        envelope = {
            "gate": "test_classification",
            "mode": "strict" if args.strict else "dry-run",
            "status": "fail" if violations else "pass",
            "indicators": {"violations": len(violations)},
            "ceilings": {"violations": 0},
            "findings": [{"rule": v.rule, "file": v.file, "detail": v.detail} for v in violations],
        }
        print(json.dumps(envelope, indent=2))
        return 1 if (violations and args.strict) else 0

    if violations:
        sys.stderr.write(f"FAIL  {len(violations)} violation(s) found:\n")
        for v in violations:
            sys.stderr.write(f"  {v}\n")
        if args.strict:
            return 1
        sys.stderr.write("\n(dry-run mode; exit 0 — pass --strict to fail)\n")
        return 0

    sys.stderr.write("OK    no violations\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
