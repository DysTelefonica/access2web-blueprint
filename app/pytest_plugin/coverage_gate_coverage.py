# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Coverage-report lookup helpers for the coverage gate plugin.

Three helpers that read the in-memory pytest-cov coverage report for a
named module path: ``_file_for_module`` resolves the measured file
path by suffix match; ``_module_covered_lines`` and
``_module_total_executable`` return the per-module line sets.

Pure helpers: each takes its inputs explicitly and returns either
``None`` (unmeasured) or a ``set[int]``. The orchestrator in
``coverage_gate.py`` consumes the resolved file path and line-span to
evaluate coverage.

W44 (#488) split this out of the original ``coverage_gate_helpers.py``
together with ``coverage_gate_messages.py`` and
``coverage_gate_resolution.py`` so each file stays under the
mutation-sites ceiling.
"""

from __future__ import annotations

from typing import Any


def _file_for_module(coverage_data: Any, module_path: str) -> str | None:
    """Return the measured file path whose name ends with ``module_path``.

    ``coverage_data`` is the in-memory coverage report after the pytest
    run. The plugin matches by suffix so it works regardless of where
    pytest was invoked from.
    """
    if coverage_data is None:
        return None
    cov_data = getattr(coverage_data, "_data", None)
    if cov_data is None:
        return None
    for file_path in cov_data.measured_files():
        if module_path in file_path or file_path.endswith(module_path):
            return file_path
    return None


def _module_covered_lines(coverage_data: Any, module_path: str) -> set[int] | None:
    """Return the executed lines for the given module file, or ``None`` when not measured."""
    file_path = _file_for_module(coverage_data, module_path)
    if file_path is None:
        return None
    cov_data = coverage_data._data
    executable = cov_data.executable_lines(file_path)  # type: ignore[no-any-return]
    executed = cov_data.executed_lines(file_path) or set()
    return executable & executed  # only executable-and-executed lines


def _module_total_executable(coverage_data: Any, module_path: str) -> set[int] | None:
    """Return the executable lines for the given module file, or ``None`` when not measured."""
    file_path = _file_for_module(coverage_data, module_path)
    if file_path is None:
        return None
    cov_data = coverage_data._data
    return set(cov_data.executable_lines(file_path))  # type: ignore[no-any-return]
