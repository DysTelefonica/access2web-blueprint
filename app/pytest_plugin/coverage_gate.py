# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp DA-13 — coverage_gate.py
"""Coverage gate plugin — 100% floor on CRITICAL_HELPERS, mutate `session.exitstatus`.

Phase 0 ships only the **declaration** of `CRITICAL_HELPERS` and the
`pytest_sessionfinish` hook. The four named helpers
(`hash_password`, `verify_password`, `issue_reset_token`,
`consume_reset_token`) are implemented in PR 4 (Phase 3). Until then, the
plugin MUST tolerate their absence — it emits a WARNING, not a failure.

Hard Rule 8: `config.exitstatus = 1` does NOT change the exit code; the
contract is to mutate `session.exitstatus` from inside `pytest_sessionfinish`.
The plugin is verified end-to-end by `pytester` in PR 4 (Phase 6).

DA-2, DA-4: the helpers themselves are Argon2id PHC strings and atomic reset
tokens. Coverage must reach 100% on every branch of every helper before a
release can pass this gate.

Lanzadera MVP CRITICAL_HELPERS (orchestrator pre-resolved, 2026-08-09):
    hash_password, verify_password, issue_reset_token, consume_reset_token.

DG-11 (issue #266, slice 1): the per-target resolution and the coverage lookup
live in `coverage_gate_helpers.py` so the gate file itself stays below the
mutation-sites ceiling. The orchestrator stays here because it owns the side
effects — stdout/stderr writes and `session.exitstatus` mutation.
"""

from __future__ import annotations

from app.pytest_plugin.coverage_gate_helpers import (
    _evaluate_helper,
    _missing_helper_warning,
    _missing_module_warning,
    _resolve_helper,
)

# Symbol names that must reach 100% coverage the moment they exist. Until they
# exist, the plugin emits a WARNING — Phase 0 ships no implementation.
CRITICAL_HELPERS: tuple[str, ...] = (
    "hash_password",
    "verify_password",
    "issue_reset_token",
    "consume_reset_token",
)

# Modules under test (Phase 4+ scope; Phase 0 walks nothing yet).
TARGET_MODULES: tuple[str, ...] = (
    "app.src.modules.lanzadera.application.credential_helpers",
    "app.src.modules.lanzadera.application.auth_reset",
)


def _try_import(module_name: str):
    """Return the imported module or ``None`` when it cannot be imported.

    Missing modules are expected during Phase 0..3 and a WARNING condition,
    not a failure. Failure means the helper exists and is under-tested.
    """
    import importlib

    try:
        return importlib.import_module(module_name)
    except ImportError:
        return None


def _emit_verdicts(warnings: list[str], failures: list[str], session) -> None:
    """Print each verdict line to stderr and mutate ``session.exitstatus``.

    Hard Rule 8: the contract is to mutate ``session.exitstatus`` from inside
    the sessionfinish hook; ``config.exitstatus = 1`` does NOT change pytest's
    exit code. We pass the session through so the mutation happens here.
    """
    # Emit one-line verdicts to stderr so reviewers see them even when pytest
    # is invoked from the CI aggregator.
    import sys

    for line in warnings:
        print(line, file=sys.stderr)
    for line in failures:
        print(line, file=sys.stderr)

    if failures:
        # Hard Rule 8: mutate session.exitstatus, not config.exitstatus.
        session.exitstatus = 1


def pytest_sessionfinish(session, exitstatus) -> None:
    """Phase 0 hook — mutate `session.exitstatus` if any CRITICAL_HELPER is under-covered.

    Behaviour:
        * Helper missing everywhere           -> emit WARNING, exit code unchanged.
        * Helper present and 100% covered     -> emit OK, exit code unchanged.
        * Helper present and <100% covered   -> emit FAIL, mutate exitstatus.

    The plugin is intentionally tolerant on Phase 0: the four helpers do not
    exist yet, so the run ends with four WARNINGS and exit code 0 (assuming
    no other test failed). Phase 4+ will fail loudly until each helper
    reaches 100%.

    DG-11 (issue #266): the per-target resolution and the coverage lookup
    moved to `coverage_gate_helpers.py` so this file stays under the
    mutation-sites ceiling.
    """
    coverage_data = getattr(session.config, "_coverage", None)
    if coverage_data is None:
        return  # coverage not enabled; pytest-cov owns the global floor (85%)

    warnings: list[str] = []
    failures: list[str] = []

    for module_name in TARGET_MODULES:
        module = _try_import(module_name)
        if module is None:
            warnings.append(_missing_module_warning(module_name))
            continue

        for helper_name in CRITICAL_HELPERS:
            file_path, error_tag = _resolve_helper(module, module_name, helper_name)
            if error_tag == "missing_helper":
                warnings.append(_missing_helper_warning(module_name, helper_name))
                continue
            if error_tag == "no_code":
                continue
            warning, failure = _evaluate_helper(file_path, module_name, helper_name, coverage_data)
            if warning is not None:
                warnings.append(warning)
            if failure is not None:
                failures.append(failure)

    _emit_verdicts(warnings, failures, session)
