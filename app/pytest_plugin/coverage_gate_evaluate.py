# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Orchestration + per-target evaluation for the coverage gate plugin.

Two functions that drive the per-CRITICAL_TARGET loop:

- ``_resolve_target`` maps a (module, owner, helper) triple to its
  source file + executable line span, or to an error tag
  (``missing_helper`` / ``no_code``).
- ``_enforce_critical_coverage`` runs the loop, accumulates warnings
  and failures, and delegates the verdict emission back to the
  orchestrator in ``coverage_gate.py``.

W45 (#490) split this out of ``coverage_gate.py`` so the plugin's
entry-point file (the pytest hooks, fixture, and StashKey) stays
under the mutation-sites ceiling. The orchestrator function does
the side-effecting work (the print-to-stderr loop and the
``session.exitstatus = 1`` mutation required by Hard Rule 8), so this
module stays testable in isolation.
"""

from __future__ import annotations

from typing import Any

from app.pytest_plugin.coverage_gate_messages import (
    _missing_helper_warning,
    _missing_module_warning,
)
from app.pytest_plugin.coverage_gate_resolution import (
    _callable_line_span,
    _evaluate_helper,
    _resolve_helper,
)


def _resolve_target(
    module: Any,
    module_name: str,
    owner_name: str | None,
    helper_name: str,
) -> tuple[str, str | None, set[int] | None, str | None]:
    owner = module if owner_name is None else getattr(module, owner_name, None)
    target_owner_name = ".".join(filter(None, (module_name, owner_name)))
    if owner is None:
        return target_owner_name, None, None, "missing_helper"
    file_path, error_tag = _resolve_helper(owner, target_owner_name, helper_name)
    line_span = _callable_line_span(owner, helper_name)
    return target_owner_name, file_path, line_span, error_tag


def _enforce_critical_coverage(
    session: Any,
    coverage_data: Any,
    critical_targets: Any,
) -> tuple[list[str], list[str]]:
    """Iterate ``critical_targets``; return ``(warnings, failures)`` lists.

    Behaviour:
        * Helper missing everywhere           -> emit WARNING, exit code unchanged.
        * Helper present and 100% covered     -> emit OK, exit code unchanged.
        * Helper present and <100% covered   -> emit FAIL, mutate exitstatus.

    The plugin is intentionally tolerant on Phase 0: the four helpers do not
    exist yet, so the run ends with four WARNINGS and exit code 0 (assuming
    no other test failed). Phase 4+ will fail loudly until each helper
    reaches 100%.

    The caller (the pytest hook in ``coverage_gate.py``) owns the
    ``session.exitstatus = 1`` mutation required by Hard Rule 8; this
    module only classifies each target into a warning or failure.
    """
    if coverage_data is None:
        return [], []

    warnings: list[str] = []
    failures: list[str] = []

    for module_name, owner_name, helper_name in critical_targets:
        from app.pytest_plugin.coverage_gate import _try_import

        module = _try_import(module_name)
        if module is None:
            warnings.append(_missing_module_warning(module_name))
            continue

        target_owner_name, file_path, line_span, error_tag = _resolve_target(
            module,
            module_name,
            owner_name,
            helper_name,
        )
        if error_tag == "missing_helper":
            warnings.append(_missing_helper_warning(target_owner_name, helper_name))
            continue
        if error_tag == "no_code" or file_path is None or line_span is None:
            continue
        warning, failure = _evaluate_helper(
            file_path,
            target_owner_name,
            helper_name,
            coverage_data,
            line_span,
        )
        if warning is not None:
            warnings.append(warning)
        if failure is not None:
            failures.append(failure)

    return warnings, failures
