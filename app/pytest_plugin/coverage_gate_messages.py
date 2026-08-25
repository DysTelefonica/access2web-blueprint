# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Verdict message builders for the coverage gate plugin.

Pure string builders; each returns a single ``str`` consumed by the
orchestrator in ``coverage_gate.py``. Kept off ``coverage_gate.py`` so
the orchestrator's mutation-site count stays bounded; kept off
``coverage_gate_helpers.py`` so the helpers file stays bounded.

W44 (#488) split this out of the original ``coverage_gate_helpers.py``
together with ``coverage_gate_coverage.py`` and
``coverage_gate_resolution.py``.
"""

from __future__ import annotations


def _missing_module_warning(module_name: str) -> str:
    """Phase 0..3 message: the whole helper module is absent."""
    return (
        f"coverage_gate WARNING: helper module '{module_name}' is not importable; "
        f"the CRITICAL_HELPERS it would host cannot be measured yet"
    )


def _missing_helper_warning(module_name: str, helper_name: str) -> str:
    """Phase 0..3 message: the helper is absent inside an existing module."""
    return f"coverage_gate WARNING: '{module_name}.{helper_name}' is not defined yet"


def _not_measured_warning(module_name: str, helper_name: str, file_path: str) -> str:
    """Helper exists but pytest-cov did not measure its file."""
    return (
        f"coverage_gate WARNING: '{module_name}.{helper_name}' exists but its "
        f"file '{file_path}' is not measured by pytest-cov"
    )


def _under_coverage_failure(module_name: str, helper_name: str, missing: set[int]) -> str:
    """Helper exists with at least one uncovered executable statement."""
    return (
        f"coverage_gate FAIL: '{module_name}.{helper_name}' is missing "
        f"{len(missing)} executed line(s): {sorted(missing)[:5]}..."
    )
