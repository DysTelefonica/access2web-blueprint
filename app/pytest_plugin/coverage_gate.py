# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp DA-13 — coverage_gate.py
"""Coverage gate plugin — 100% floor on CRITICAL_HELPERS, mutate `session.exitstatus`.

The plugin captures pytest-cov's supported ``cov`` fixture and evaluates four
explicit ``CRITICAL_TARGETS`` inside a try-last ``pytest_runtestloop`` wrapper.
Each live callable is checked against its own executable source-line span;
undercoverage increments ``session.testsfailed`` and sets ``session.exitstatus``
to 1 before pytest-cov replaces its active coverage object.

Class methods carry their owner explicitly so ``CredentialHasherArgon2id.hash``
and ``CredentialHasherArgon2id.verify`` resolve without module-level wrappers.

Hard Rule 8: ``config.exitstatus = 1`` does NOT change the exit code. The
runtest-loop wrapper mutates ``session.testsfailed`` before final exit handling.

DA-2, DA-4: the helpers themselves are Argon2id PHC strings and atomic reset
tokens. Coverage must reach 100 % of each helper's executable statements before a
release can pass this gate.

Lanzadera MVP CRITICAL_HELPERS (W-series current):
    hash, verify  # methods on CredentialHasherArgon2id (W07 #434)
    issue_reset_token, consume_reset_token  # domain services (D90)

DG-11 (issue #266, slice 1): the per-target resolution and the coverage lookup
live in ``coverage_gate_helpers.py`` so the gate file itself stays below the
mutation-sites ceiling. The orchestrator stays here because it owns the side
effects — stdout/stderr writes and ``session.exitstatus`` mutation.
"""

from __future__ import annotations

from typing import Any

import pytest

from app.pytest_plugin.coverage_gate_evaluate import (
    _enforce_critical_coverage as _evaluate_coverage,
)

# Exact (module, optional owner, helper) targets required by DA-2 and DA-4.
CRITICAL_TARGETS: tuple[tuple[str, str | None, str], ...] = (
    (
        "app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id",
        "CredentialHasherArgon2id",
        "hash",
    ),
    (
        "app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id",
        "CredentialHasherArgon2id",
        "verify",
    ),
    (
        "app.src.modules.lanzadera.domain.services.issue_reset_token",
        None,
        "issue_reset_token",
    ),
    (
        "app.src.modules.lanzadera.domain.services.consume_reset_token",
        None,
        "consume_reset_token",
    ),
)
CRITICAL_HELPERS: tuple[str, ...] = tuple(target[2] for target in CRITICAL_TARGETS)

# Modules under test (Phase 4+ scope; Phase 0 walks nothing yet).
# The real modules live in ``adapters/crypto`` and ``domain/services``;
# the older ``application.{credential_helpers,auth_reset}`` paths the
# plugin declared at design time never materialised.
TARGET_MODULES: tuple[str, ...] = tuple(
    dict.fromkeys(module_name for module_name, _, _ in CRITICAL_TARGETS)
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


def _enforce_critical_coverage(session: Any, coverage_data: Any) -> None:
    """Backward-compat wrapper — delegates to ``coverage_gate_evaluate``.

    W45 extracted the orchestrator into ``coverage_gate_evaluate.py`` for
    the mutation-sites ratchet. This shim preserves the original
    ``(session, coverage_data)`` signature so existing tests continue
    to work without changes. The shim imports the implementation lazily
    so monkey-patching ``coverage_gate._try_import`` propagates to the
    implementation (the implementation reads ``coverage_gate._try_import``
    via lazy import per iteration).
    """
    warnings, failures = _evaluate_coverage(session, coverage_data, CRITICAL_TARGETS)
    _emit_verdicts(warnings, failures, session)


def _emit_verdicts(warnings: list[str], failures: list[str], session: Any) -> None:
    """Print each verdict line to stderr and mutate ``session.exitstatus``.

    Hard Rule 8: the contract is to mutate ``session.exitstatus`` from inside
    the sessionfinish hook; ``config.exitstatus = 1`` does NOT change pytest's
    exit code. We pass the session through so the mutation happens here.
    """
    import sys

    for line in warnings:
        print(line, file=sys.stderr)
    for line in failures:
        print(line, file=sys.stderr)

    if failures:
        session.exitstatus = 1
        if hasattr(session, "testsfailed"):
            session.testsfailed += 1


_COVERAGE_KEY = pytest.StashKey[Any]()


@pytest.fixture(autouse=True)
def capture_pytest_cov(cov, request) -> None:
    request.config.stash[_COVERAGE_KEY] = cov


@pytest.hookimpl(wrapper=True, trylast=True)
def pytest_runtestloop(session):
    result = yield
    coverage_data = session.config.stash.get(_COVERAGE_KEY, None)
    _enforce_critical_coverage(session, coverage_data)
    return result
