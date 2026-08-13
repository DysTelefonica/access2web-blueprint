# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp DA-13 — coverage_gate_helpers.py
"""Pure helpers for the coverage gate plugin.

These are kept off `coverage_gate.py` so the gate file itself stays under the
mutation-sites ceiling (DG-11, issue #266): every helper extracted into this
module would otherwise inflate the per-file atom count and exceed the BASELINE
the gate established at `sites=118, target=100`.

The helpers are pure: each takes its inputs explicitly and returns either a
string verdict message or a `(warning, failure)` tuple. The orchestrator in
`coverage_gate.py` keeps the side-effecting work — the print-to-stderr loop and
the `session.exitstatus = 1` mutation required by Hard Rule 8 — so this module
stays testable in isolation and pytest does not auto-load it as a plugin.
"""

from __future__ import annotations


# --------------------------------------------------------------------------------------------
# Coverage lookup
# --------------------------------------------------------------------------------------------

def _module_covered_lines(coverage_data, module_path: str) -> set[int] | None:
    """Return the executed lines for the given module file, or ``None`` when not measured.

    coverage_data is the in-memory coverage report after the pytest run. Keys
    are the relative paths of measured files; we match by suffix so the
    plugin works regardless of where pytest was invoked from.
    """
    if coverage_data is None:
        return None
    cov_data = getattr(coverage_data, "_data", None)
    if cov_data is None:
        return None
    measured_files = cov_data.measured_files()
    for file_path in measured_files:
        if module_path in file_path or file_path.endswith(module_path):
            executable = cov_data.executable_lines(file_path)
            executed = cov_data.executed_lines(file_path) or set()
            return executable & executed  # only executable-and-executed lines
    return None


def _module_total_executable(coverage_data, module_path: str) -> set[int] | None:
    """Return the executable lines for the given module file, or ``None`` when not measured."""
    if coverage_data is None:
        return None
    cov_data = getattr(coverage_data, "_data", None)
    if cov_data is None:
        return None
    measured_files = cov_data.measured_files()
    for file_path in measured_files:
        if module_path in file_path or file_path.endswith(module_path):
            return set(cov_data.executable_lines(file_path))
    return None


# --------------------------------------------------------------------------------------------
# Verdict messages — pure string builders
# --------------------------------------------------------------------------------------------

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


def _under_coverage_failure(
    module_name: str, helper_name: str, missing: set[int]
) -> str:
    """Helper exists and is below 100% branch coverage on at least one line."""
    return (
        f"coverage_gate FAIL: '{module_name}.{helper_name}' is missing "
        f"{len(missing)} executed line(s): {sorted(missing)[:5]}..."
    )


# --------------------------------------------------------------------------------------------
# Per-target resolution and evaluation
# --------------------------------------------------------------------------------------------

def _resolve_helper(
    module, module_name: str, helper_name: str
) -> tuple[str | None, str | None]:
    """Resolve one (module, helper) pair to its source file or an error tag.

    The caller passes the already-imported module so we do not re-import per
    helper. Returns ``(file_path, None)`` on success; otherwise
    ``(None, "missing_helper")`` when the attribute is absent and
    ``(None, "no_code")`` when the helper has no ``__code__`` object.
    """
    helper = getattr(module, helper_name, None)
    if helper is None:
        return None, "missing_helper"
    helper_file = getattr(helper, "__code__", None)
    if helper_file is None:
        return None, "no_code"
    return helper_file.co_filename, None


def _evaluate_helper(
    file_path: str,
    module_name: str,
    helper_name: str,
    coverage_data,
) -> tuple[str | None, str | None]:
    """Return ``(warning, failure)`` for one present helper's coverage check.

    Both are ``None`` when the helper is fully covered. ``warning`` carries a
    "not measured" message when pytest-cov did not see the file; ``failure``
    carries a "missing lines" message when the file was measured but under
    100%.
    """
    executed = _module_covered_lines(coverage_data, file_path)
    executable = _module_total_executable(coverage_data, file_path)
    if executed is None or executable is None:
        return _not_measured_warning(module_name, helper_name, file_path), None
    missing = executable - executed
    if missing:
        return None, _under_coverage_failure(module_name, helper_name, missing)
    return None, None
