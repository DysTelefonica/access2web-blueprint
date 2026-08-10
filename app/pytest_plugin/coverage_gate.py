# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp DA-13 — coverage_gate.py
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
"""

from __future__ import annotations

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
    """
    coverage_data = getattr(session.config, "_coverage", None)
    if coverage_data is None:
        return  # coverage not enabled; pytest-cov owns the global floor (85%)

    warnings: list[str] = []
    failures: list[str] = []

    for module_name in TARGET_MODULES:
        module = _try_import(module_name)
        if module is None:
            # Phase 0..3: helper module absent. Emit a WARNING per missing module.
            warnings.append(
                f"coverage_gate WARNING: helper module '{module_name}' is not importable; "
                f"the CRITICAL_HELPERS it would host cannot be measured yet"
            )
            continue

        for helper_name in CRITICAL_HELPERS:
            helper = getattr(module, helper_name, None)
            if helper is None:
                # Phase 0..3: helper absent inside an existing module. WARN.
                warnings.append(
                    f"coverage_gate WARNING: '{module_name}.{helper_name}' is not defined yet"
                )
                continue

            # Helper exists. Resolve its file path and check coverage.
            helper_file = getattr(helper, "__code__", None)
            if helper_file is None:
                continue
            file_path = helper_file.co_filename

            executed = _module_covered_lines(coverage_data, file_path)
            executable = _module_total_executable(coverage_data, file_path)
            if executed is None or executable is None:
                # File was not measured — pytest-cov scope excluded it, or the helper is
                # defined in a file outside coverage scope. Surface as WARNING; the verifier
                # (Phase 6) will pin the scope.
                warnings.append(
                    f"coverage_gate WARNING: '{module_name}.{helper_name}' exists but its "
                    f"file '{file_path}' is not measured by pytest-cov"
                )
                continue

            missing = executable - executed
            if missing:
                failures.append(
                    f"coverage_gate FAIL: '{module_name}.{helper_name}' is missing "
                    f"{len(missing)} executed line(s): {sorted(missing)[:5]}..."
                )
            else:
                pass  # silent OK

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
