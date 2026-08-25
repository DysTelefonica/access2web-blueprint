# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Per-target resolution and coverage evaluation for the coverage gate plugin.

Three helpers that drive the orchestrator's per-target loop:

- ``_resolve_helper`` maps a (module, helper) pair to its source file
  or an error tag (``missing_helper`` / ``no_code``).
- ``_callable_line_span`` returns the executable line span of the
  resolved helper via ``inspect.getsourcelines``.
- ``_evaluate_helper`` consumes the resolved file path + line span and
  the in-memory coverage report and returns the ``(warning,
  failure)`` verdict pair.

Pure helpers: each takes its inputs explicitly and returns either
strings or tuples. The orchestrator in ``coverage_gate.py`` keeps the
side-effecting work (the print-to-stderr loop and the ``session.exitstatus``
mutation required by Hard Rule 8), so this module stays testable in
isolation.

W44 (#488) split this out of the original ``coverage_gate_helpers.py``
together with ``coverage_gate_messages.py`` and
``coverage_gate_coverage.py`` so each file stays under the
mutation-sites ceiling.
"""

from __future__ import annotations

import inspect
from pathlib import Path

from app.pytest_plugin.coverage_gate_messages import (
    _not_measured_warning,
    _under_coverage_failure,
)


def _resolve_helper(module, module_name: str, helper_name: str) -> tuple[str | None, str | None]:
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


def _callable_line_span(owner, helper_name: str) -> set[int] | None:
    helper = getattr(owner, helper_name, None)
    if helper is None:
        return None
    try:
        source, start = inspect.getsourcelines(inspect.unwrap(helper))
    except (OSError, TypeError):
        return None
    return set(range(start, start + len(source)))


def _evaluate_helper(
    file_path: str,
    module_name: str,
    helper_name: str,
    coverage_data,
    line_span: set[int],
) -> tuple[str | None, str | None]:
    """Return ``(warning, failure)`` for one present helper's coverage check.

    Both are ``None`` when the helper is fully covered. ``warning`` carries a
    "not measured" message when pytest-cov did not see the file; ``failure``
    carries a "missing lines" message when the file was measured but under
    100%.
    """
    expected_path = Path(file_path).resolve()
    measured_file = next(
        (
            path
            for path in coverage_data.get_data().measured_files()
            if Path(path).resolve() == expected_path
        ),
        None,
    )
    if measured_file is None:
        return _not_measured_warning(module_name, helper_name, file_path), None
    _, executable, _, missing_lines, _ = coverage_data.analysis2(measured_file)
    missing = set(missing_lines) & set(executable) & line_span
    if missing:
        return None, _under_coverage_failure(module_name, helper_name, missing)
    return None, None
