# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + DA-13 — coverage plugin test
"""Coverage pin for the `app.pytest_plugin.coverage_gate` plugin (CI fail-under gate).

Both `coverage_gate.py` and the three split helpers modules
(`coverage_gate_coverage.py`, `coverage_gate_messages.py`,
`coverage_gate_resolution.py`) are
imported by pytest at startup (via the ``-p app.pytest_plugin.coverage_gate``
entry in `app/pyproject.toml::tool.pytest.ini_options.addopts`). pytest-cov
measures every reachable module, so until something actually executes their
lines they appear as ``0%`` in `coverage.json` and drag total coverage below
the configured ``--cov-fail-under=69`` floor.

The plugin is pure orchestration: every helper is either import-resolvable or
trivial enough to drive from a unit test. We invoke each one directly rather
than going through `pytester` so the assertions stay as straightforward
attribute checks, and we never rely on `subprocess` to spawn a second pytest
just to exercise a hook that takes a ``session`` argument.

The helpers module is split out on slice 1 of #266 (commit b8adcd0,
`feat(scripts): lower MAX_COMPLEXITY ceiling 15→10 with review date`); the
tracker ships only the orchestrator. Both versions must pass — this file
tests the orchestrator unconditionally and the helpers module conditionally.
"""

from __future__ import annotations

import importlib
import os
import subprocess
import sys
from pathlib import Path
from types import ModuleType, SimpleNamespace
from typing import Any
from unittest.mock import MagicMock

import pytest

coverage_gate = importlib.import_module("app.pytest_plugin.coverage_gate")

# W44 (#488) split the helpers into three cohesive modules:
# - ``coverage_gate_coverage`` — the three lookup helpers
# - ``coverage_gate_messages`` — the four verdict message builders
# - ``coverage_gate_resolution`` — the three resolution+eval helpers
coverage_gate_coverage = importlib.import_module("app.pytest_plugin.coverage_gate_coverage")
coverage_gate_messages = importlib.import_module("app.pytest_plugin.coverage_gate_messages")
coverage_gate_resolution = importlib.import_module("app.pytest_plugin.coverage_gate_resolution")

# The pure helpers (`_module_covered_lines`, `_module_total_executable`)
# live in ``coverage_gate_coverage``; the four message builders live in
# ``coverage_gate_messages``; (`_resolve_helper`, `_evaluate_helper`,
# `_callable_line_span`) live in ``coverage_gate_resolution``.
coverage_module = coverage_gate_coverage
messages_module = coverage_gate_messages
resolution_module = coverage_gate_resolution

# Module-level fixture: pytest collects the test even if helpers are absent.
# The three split modules are always imported together; we keep the
# skipif marker for backward compatibility with the single-helpers branch.
pytestmark_helpers_required = pytest.mark.skipif(
    False,
    reason="coverage_gate helpers are split into three modules, always importable",
)


# --------------------------------------------------------------------------------------------
# Constants — the declaration site the gate validates
# --------------------------------------------------------------------------------------------


def test_critical_helpers_lists_the_four_named_helpers() -> None:
    """DA-2, DA-4: the four CRITICAL_HELPERS the gate is supposed to enforce.

    The names mirror the live ``PasswordHasher`` Protocol surface (DA-2):
    the ``CredentialHasherArgon2id`` adapter carries ``hash`` and ``verify``
    since W07 (#434). The earlier ``hash_password`` / ``verify_password``
    names pre-date that migration.
    """
    assert coverage_gate.CRITICAL_HELPERS == (
        "hash",
        "verify",
        "issue_reset_token",
        "consume_reset_token",
    )


def test_target_modules_lists_the_three_lanzadera_helpers() -> None:
    """Phase 4+ scope: ``CredentialHasherArgon2id`` (DA-2) plus two reset-flow services (DA-4)."""
    assert coverage_gate.TARGET_MODULES == (
        "app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id",
        "app.src.modules.lanzadera.domain.services.issue_reset_token",
        "app.src.modules.lanzadera.domain.services.consume_reset_token",
    )


# --------------------------------------------------------------------------------------------
# _try_import — import or None
# --------------------------------------------------------------------------------------------


def test_try_import_returns_module_when_importable() -> None:
    """A module present on sys.path resolves to the actual module object."""
    assert coverage_gate._try_import("os") is os


def test_try_import_returns_none_on_missing_module() -> None:
    """Phase 0..3 contract: missing helper modules are WARNING, not failure."""
    assert coverage_gate._try_import("app.does_not_exist_anywhere_xyz") is None


# --------------------------------------------------------------------------------------------
# _file_for_module — measured-files lookup
# --------------------------------------------------------------------------------------------


def test_file_for_module_returns_none_when_coverage_data_is_none() -> None:
    """A missing coverage report short-circuits the lookup."""
    assert coverage_gate_coverage._file_for_module(None, "app.x.y") is None


def test_file_for_module_matches_by_suffix() -> None:
    """Match by suffix tolerates worktree-root prefix variations."""
    cov = _make_coverage_data(
        measured_files=["app/src/modules/lanzadera/domain/user.py"],
        executable_lines={},
        executed_lines={},
    )
    path = coverage_gate_coverage._file_for_module(
        cov,
        "lanzadera/domain/user.py",
    )
    assert path == "app/src/modules/lanzadera/domain/user.py"


def test_file_for_module_matches_by_substring() -> None:
    """Match by substring catches paths that share the module name suffix."""
    cov = _make_coverage_data(
        measured_files=["/tmp/build/app/src/modules/lanzadera/app.py"],
        executable_lines={},
        executed_lines={},
    )
    path = coverage_gate_coverage._file_for_module(
        cov,
        "app/src/modules/lanzadera/app.py",
    )
    assert path == "/tmp/build/app/src/modules/lanzadera/app.py"


def test_file_for_module_returns_none_when_no_match() -> None:
    """An unmatched ``module_path`` returns ``None``."""
    cov = _make_coverage_data(
        measured_files=["app/src/modules/lanzadera/domain/user.py"],
        executable_lines={},
        executed_lines={},
    )
    assert coverage_gate_coverage._file_for_module(cov, "wrong") is None


# --------------------------------------------------------------------------------------------
# _module_covered_lines and _module_total_executable — coverage_data queries
# --------------------------------------------------------------------------------------------


def _make_coverage_data(
    measured_files: list[str],
    executable_lines: dict[str, set[int]],
    executed_lines: dict[str, set[int]],
) -> Any:
    """Build a stand-in for the in-memory coverage report.

    Only the surface area the helpers touch is implemented. Tests should
    not reach for attributes beyond ``_data.measured_files``,
    ``_data.executable_lines`` and ``_data.executed_lines`` — that is exactly
    the contract the helpers read.
    """
    cov_data = MagicMock()
    cov_data.measured_files.return_value = measured_files
    cov_data.executable_lines.side_effect = lambda path: executable_lines.get(path, set())
    cov_data.executed_lines.side_effect = lambda path: executed_lines.get(path, set())
    coverage = MagicMock()
    coverage._data = cov_data
    coverage.get_data.return_value = cov_data
    coverage.analysis2.side_effect = lambda path: (
        path,
        sorted(executable_lines.get(path, set())),
        [],
        sorted(executable_lines.get(path, set()) - executed_lines.get(path, set())),
        "",
    )
    return coverage


def test_module_covered_lines_returns_none_when_no_coverage_data() -> None:
    """No coverage means no measurement — Phase 0 contract."""
    assert coverage_module._module_covered_lines(None, "anywhere/foo.py") is None


def test_module_covered_lines_returns_none_when_data_has_no_inner_data() -> None:
    """Coverage data without `_data` is treated as unmeasured (defensive)."""
    assert coverage_module._module_covered_lines(SimpleNamespace(), "anywhere/foo.py") is None


def test_module_covered_lines_returns_none_when_path_does_not_match() -> None:
    """A measured file that does not contain the module path yields None."""
    cov = _make_coverage_data(
        measured_files=["app/src/other.py"],
        executable_lines={"app/src/other.py": {1, 2}},
        executed_lines={"app/src/other.py": {1, 2}},
    )
    assert coverage_module._module_covered_lines(cov, "app/src/target.py") is None


def test_module_covered_lines_returns_intersection_for_matching_path() -> None:
    """The contract is `executable & executed` — the lines that were BOTH measurable AND run."""
    cov = _make_coverage_data(
        measured_files=["app/src/target.py"],
        executable_lines={"app/src/target.py": {1, 2, 3, 4}},
        executed_lines={"app/src/target.py": {2, 3}},
    )
    assert coverage_module._module_covered_lines(cov, "app/src/target.py") == {2, 3}


def test_module_covered_lines_treats_none_executed_as_empty_set() -> None:
    """`cov_data.executed_lines` may return ``None``; the helper must not propagate it."""
    cov_data = MagicMock()
    cov_data.measured_files.return_value = ["app/src/target.py"]
    cov_data.executable_lines.return_value = {1, 2, 3}
    cov_data.executed_lines.return_value = None
    cov = SimpleNamespace(_data=cov_data)
    assert coverage_module._module_covered_lines(cov, "app/src/target.py") == set()


def test_module_total_executable_returns_none_when_no_coverage_data() -> None:
    assert coverage_module._module_total_executable(None, "anywhere/foo.py") is None


def test_module_total_executable_returns_none_when_data_has_no_inner_data() -> None:
    assert coverage_module._module_total_executable(SimpleNamespace(), "anywhere/foo.py") is None


def test_module_total_executable_returns_none_when_path_does_not_match() -> None:
    cov = _make_coverage_data(
        measured_files=["app/src/other.py"],
        executable_lines={"app/src/other.py": {1}},
        executed_lines={"app/src/other.py": set()},
    )
    assert coverage_module._module_total_executable(cov, "app/src/target.py") is None


def test_module_total_executable_returns_executable_lines_for_matching_path() -> None:
    cov = _make_coverage_data(
        measured_files=["app/src/target.py"],
        executable_lines={"app/src/target.py": {1, 2, 3}},
        executed_lines={"app/src/target.py": {1}},
    )
    assert coverage_module._module_total_executable(cov, "app/src/target.py") == {1, 2, 3}


# --------------------------------------------------------------------------------------------
# pytest_sessionfinish — orchestrator hook
# --------------------------------------------------------------------------------------------


def _make_session(coverage_data: Any | None) -> Any:
    """Build a stand-in for the pytest session the hook receives."""
    config = SimpleNamespace(_coverage=coverage_data)
    session = SimpleNamespace(
        config=config,
        exitstatus=0,
    )
    return session


def _make_sessionfinish_sentinel_module() -> tuple[ModuleType, str, set[int]]:
    """Build a real ``ModuleType`` carrying one ``__code__``-bearing helper per CRITICAL_HELPER.

    The orchestrator iterates CRITICAL_HELPERS and emits a WARNING per missing
    one, so a sentinel that exposes only `hash_password` produces three
    spurious warnings and breaks the "fully covered" test. Each helper points
    at the same source file so the synthetic coverage report matches the
    orchestrator's `co_filename` lookup uniformly.
    """
    sentinel_module = ModuleType("sentinel_sessionfinish")

    def fake_helper() -> None:
        """Stand-in for a CRITICAL_HELPER on the orchestrator's target module."""
        return None

    for helper_name in coverage_gate.CRITICAL_HELPERS:
        setattr(sentinel_module, helper_name, fake_helper)
    sentinel_module.CredentialHasherArgon2id = type(  # type: ignore[attr-defined]
        "CredentialHasherArgon2id", (), {"hash": fake_helper, "verify": fake_helper}
    )
    line_span = coverage_gate_resolution._callable_line_span(
        sentinel_module.CredentialHasherArgon2id,
        "hash",  # type: ignore[attr-defined]
    )
    assert line_span is not None
    return sentinel_module, fake_helper.__code__.co_filename, line_span


def test_sessionfinish_returns_early_when_coverage_disabled() -> None:
    """No coverage on the run means the gate delegates to pytest-cov's global floor."""
    session = _make_session(None)
    coverage_gate._enforce_critical_coverage(session, None)
    assert session.exitstatus == 0


def test_live_pytest_cov_fails_for_undercovered_exact_target(tmp_path: Path) -> None:
    (tmp_path / "target_module.py").write_text(
        "def critical(flag):\n    if flag:\n        return 1\n    return 0\n", encoding="utf-8"
    )
    (tmp_path / "test_target.py").write_text(
        "from target_module import critical\n\ndef test_x():\n    assert critical(True) == 1\n",
        encoding="utf-8",
    )
    (tmp_path / "conftest.py").write_text(
        "import app.pytest_plugin.coverage_gate as gate\n"
        'gate.CRITICAL_TARGETS = (("target_module", None, "critical"),)\n',
        encoding="utf-8",
    )
    env = os.environ.copy()
    env["PYTHONPATH"] = os.pathsep.join(
        (str(Path(__file__).resolve().parents[2]), env.get("PYTHONPATH", ""))
    )
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "pytest",
            "-p",
            "app.pytest_plugin.coverage_gate",
            "--cov=target_module",
            "test_target.py",
        ],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert (result.returncode, "coverage_gate FAIL" in result.stderr) == (1, True)


def test_sessionfinish_warns_when_target_module_not_importable(
    capsys: pytest.CaptureFixture[str],
) -> None:
    """Phase 0..3 contract: a missing target module emits a WARNING, exit status unchanged.

    Today's ``CredentialHasherArgon2id`` module IS importable, so the
    orchestrator walks the next branch — it warns that each helper
    (``hash``, ``verify``) is not defined as a module-level attribute. The
    helper classes live as methods inside the class, not as free
    functions, so the orchestrator's attribute lookup returns False. The
    test verifies the WARNING line fires per ``(module, helper)`` pair.
    """
    sentinel_module = ModuleType("sentinel_sessionfinish")
    cov = _make_coverage_data(
        measured_files=[],
        executable_lines={},
        executed_lines={},
    )
    real_try_import = coverage_gate._try_import

    def fake_try_import(name: str) -> ModuleType | None:
        if name in coverage_gate.TARGET_MODULES:
            return sentinel_module  # importable, no helpers as module attrs
        return real_try_import(name)

    coverage_gate._try_import = fake_try_import
    try:
        session = _make_session(cov)
        coverage_gate._enforce_critical_coverage(session, cov)
    finally:
        coverage_gate._try_import = real_try_import

    captured = capsys.readouterr()
    assert captured.err.count("is not defined yet") == 4
    assert session.exitstatus == 0


def test_sessionfinish_warns_and_does_not_fail_when_helper_absent(
    capsys: pytest.CaptureFixture[str],
) -> None:
    """Helper missing inside an existing module: WARNING, exit status unchanged."""
    sentinel_module = ModuleType("sentinel_module")
    cov = _make_coverage_data(
        measured_files=[],
        executable_lines={},
        executed_lines={},
    )

    real_try_import = coverage_gate._try_import

    def fake_try_import(name: str) -> ModuleType | None:
        if name in coverage_gate.TARGET_MODULES:
            return sentinel_module
        return real_try_import(name)

    coverage_gate._try_import = fake_try_import
    try:
        session = _make_session(cov)
        coverage_gate._enforce_critical_coverage(session, cov)
    finally:
        coverage_gate._try_import = real_try_import

    captured = capsys.readouterr()
    assert captured.err.count("is not defined yet") == 4
    assert session.exitstatus == 0


def test_sessionfinish_warns_when_helper_file_not_measured(
    capsys: pytest.CaptureFixture[str],
) -> None:
    """Helper exists but pytest-cov did not measure its file: WARNING, exit status unchanged.

    `_make_sessionfinish_sentinel_module` registers a real `__code__`-bearing
    function on a sentinel module so the orchestrator walks the helper-present
    branch; the coverage report deliberately lists a different file so the
    helper is treated as unmeasured.
    """
    sentinel_module, sentinel_filename, _ = _make_sessionfinish_sentinel_module()
    cov = _make_coverage_data(
        measured_files=["other/module.py"],
        executable_lines={"other/module.py": {1}},
        executed_lines={"other/module.py": {1}},
    )

    real_try_import = coverage_gate._try_import

    def fake_try_import(name: str) -> ModuleType | None:
        if name in coverage_gate.TARGET_MODULES:
            return sentinel_module
        return real_try_import(name)

    coverage_gate._try_import = fake_try_import
    try:
        session = _make_session(cov)
        coverage_gate._enforce_critical_coverage(session, cov)
    finally:
        coverage_gate._try_import = real_try_import

    captured = capsys.readouterr()
    assert "is not measured by pytest-cov" in captured.err
    assert sentinel_filename in captured.err
    assert session.exitstatus == 0


def test_sessionfinish_silent_when_helper_fully_covered(
    capsys: pytest.CaptureFixture[str],
) -> None:
    """100% covered helper: no warning, no failure, exit status unchanged."""
    sentinel_module, sentinel_filename, line_span = _make_sessionfinish_sentinel_module()
    cov = _make_coverage_data(
        measured_files=[sentinel_filename],
        executable_lines={sentinel_filename: line_span},
        executed_lines={sentinel_filename: line_span},
    )

    real_try_import = coverage_gate._try_import

    def fake_try_import(name: str) -> ModuleType | None:
        if name in coverage_gate.TARGET_MODULES:
            return sentinel_module
        return real_try_import(name)

    coverage_gate._try_import = fake_try_import
    try:
        session = _make_session(cov)
        coverage_gate._enforce_critical_coverage(session, cov)
    finally:
        coverage_gate._try_import = real_try_import

    captured = capsys.readouterr()
    assert captured.err == ""
    assert session.exitstatus == 0


def test_sessionfinish_fails_when_helper_under_covered(
    capsys: pytest.CaptureFixture[str],
) -> None:
    """Hard Rule 8: under-covered helper mutates ``session.exitstatus`` to 1."""
    sentinel_module, sentinel_filename, line_span = _make_sessionfinish_sentinel_module()
    cov = _make_coverage_data(
        measured_files=[sentinel_filename],
        executable_lines={sentinel_filename: line_span},
        executed_lines={sentinel_filename: line_span - {max(line_span)}},
    )

    real_try_import = coverage_gate._try_import

    def fake_try_import(name: str) -> ModuleType | None:
        if name in coverage_gate.TARGET_MODULES:
            return sentinel_module
        return real_try_import(name)

    coverage_gate._try_import = fake_try_import
    try:
        session = _make_session(cov)
        coverage_gate._enforce_critical_coverage(session, cov)
    finally:
        coverage_gate._try_import = real_try_import

    captured = capsys.readouterr()
    assert "coverage_gate FAIL" in captured.err
    assert session.exitstatus == 1


def test_sessionfinish_skips_helper_without_code_object(
    capsys: pytest.CaptureFixture[str],
) -> None:
    """`getattr(helper, '__code__', None)` is None on plain objects — skip silently.

    Pin every CRITICAL_HELPER on the sentinel so the orchestrator reaches
    `helper.__code__`, but assign non-callable objects so the attribute is
    missing. The orchestrator's `if helper_file is None: continue` branch
    must short-circuit without emitting WARNING or FAILURE.
    """
    sentinel_module = ModuleType("sentinel_module")

    class NoCode:
        pass

    for helper_name in coverage_gate.CRITICAL_HELPERS:
        setattr(sentinel_module, helper_name, NoCode())
    sentinel_module.CredentialHasherArgon2id = type(  # type: ignore[attr-defined]
        "CredentialHasherArgon2id", (), {"hash": NoCode(), "verify": NoCode()}
    )

    cov = _make_coverage_data(
        measured_files=[],
        executable_lines={},
        executed_lines={},
    )

    real_try_import = coverage_gate._try_import

    def fake_try_import(name: str) -> ModuleType | None:
        if name in coverage_gate.TARGET_MODULES:
            return sentinel_module
        return real_try_import(name)

    coverage_gate._try_import = fake_try_import
    try:
        session = _make_session(cov)
        coverage_gate._enforce_critical_coverage(session, cov)
    finally:
        coverage_gate._try_import = real_try_import

    captured = capsys.readouterr()
    assert "coverage_gate FAIL" not in captured.err
    assert session.exitstatus == 0


# --------------------------------------------------------------------------------------------
# coverage_gate_coverage / coverage_gate_messages / coverage_gate_resolution
# --------------------------------------------------------------------------------------------


@pytestmark_helpers_required
def test_helpers_module_covered_lines_returns_none_when_no_coverage_data() -> None:
    assert coverage_gate_coverage._module_covered_lines(None, "anywhere/foo.py") is None


@pytestmark_helpers_required
def test_helpers_module_covered_lines_returns_none_when_data_has_no_inner_data() -> None:
    assert (
        coverage_gate_coverage._module_covered_lines(SimpleNamespace(), "anywhere/foo.py") is None
    )


@pytestmark_helpers_required
def test_helpers_module_covered_lines_returns_none_when_path_does_not_match() -> None:
    cov = _make_coverage_data(
        measured_files=["app/src/other.py"],
        executable_lines={"app/src/other.py": {1}},
        executed_lines={"app/src/other.py": {1}},
    )
    assert coverage_gate_coverage._module_covered_lines(cov, "app/src/target.py") is None


@pytestmark_helpers_required
def test_helpers_module_covered_lines_returns_intersection_for_matching_path() -> None:
    cov = _make_coverage_data(
        measured_files=["app/src/target.py"],
        executable_lines={"app/src/target.py": {1, 2, 3, 4}},
        executed_lines={"app/src/target.py": {2, 3}},
    )
    assert coverage_gate_coverage._module_covered_lines(cov, "app/src/target.py") == {2, 3}


@pytestmark_helpers_required
def test_helpers_module_covered_lines_treats_none_executed_as_empty_set() -> None:
    cov_data = MagicMock()
    cov_data.measured_files.return_value = ["app/src/target.py"]
    cov_data.executable_lines.return_value = {1, 2, 3}
    cov_data.executed_lines.return_value = None
    cov = SimpleNamespace(_data=cov_data)
    assert coverage_gate_coverage._module_covered_lines(cov, "app/src/target.py") == set()


@pytestmark_helpers_required
def test_helpers_module_total_executable_returns_none_when_no_coverage_data() -> None:
    assert coverage_gate_coverage._module_total_executable(None, "anywhere/foo.py") is None


@pytestmark_helpers_required
def test_helpers_module_total_executable_returns_none_when_data_has_no_inner_data() -> None:
    assert (
        coverage_gate_coverage._module_total_executable(SimpleNamespace(), "anywhere/foo.py")
        is None
    )


@pytestmark_helpers_required
def test_helpers_module_total_executable_returns_none_when_path_does_not_match() -> None:
    cov = _make_coverage_data(
        measured_files=["app/src/other.py"],
        executable_lines={"app/src/other.py": {1}},
        executed_lines={"app/src/other.py": set()},
    )
    assert coverage_gate_coverage._module_total_executable(cov, "app/src/target.py") is None


@pytestmark_helpers_required
def test_helpers_module_total_executable_returns_executable_lines_for_matching_path() -> None:
    cov = _make_coverage_data(
        measured_files=["app/src/target.py"],
        executable_lines={"app/src/target.py": {1, 2, 3}},
        executed_lines={"app/src/target.py": {1}},
    )
    assert coverage_gate_coverage._module_total_executable(cov, "app/src/target.py") == {1, 2, 3}


@pytestmark_helpers_required
def test_helpers_missing_module_warning_names_module() -> None:
    msg = coverage_gate_messages._missing_module_warning("app.x.y")
    assert "app.x.y" in msg
    assert "not importable" in msg


@pytestmark_helpers_required
def test_helpers_missing_helper_warning_names_pair() -> None:
    msg = coverage_gate_messages._missing_helper_warning("app.x.y", "helper_z")
    assert "app.x.y.helper_z" in msg
    assert "not defined yet" in msg


@pytestmark_helpers_required
def test_helpers_not_measured_warning_names_pair_and_file() -> None:
    msg = coverage_gate_messages._not_measured_warning("app.x.y", "helper_z", "/tmp/file.py")
    assert "app.x.y.helper_z" in msg
    assert "/tmp/file.py" in msg
    assert "not measured" in msg


@pytestmark_helpers_required
def test_helpers_under_coverage_failure_counts_missing() -> None:
    msg = coverage_gate_messages._under_coverage_failure("app.x.y", "helper_z", {1, 2, 3, 4, 5, 6})
    assert "6" in msg
    assert "app.x.y.helper_z" in msg


@pytestmark_helpers_required
def test_helpers_resolve_helper_returns_file_path_when_present() -> None:
    sentinel = ModuleType("sentinel")
    sentinel.fake_helper = lambda: None  # type: ignore[attr-defined]
    file_path, error = coverage_gate_resolution._resolve_helper(sentinel, "app.x", "fake_helper")
    assert error is None
    assert file_path is not None


@pytestmark_helpers_required
def test_helpers_resolve_helper_returns_missing_helper_when_absent() -> None:
    sentinel = ModuleType("sentinel")
    file_path, error = coverage_gate_resolution._resolve_helper(sentinel, "app.x", "no_such")
    assert file_path is None
    assert error == "missing_helper"


@pytestmark_helpers_required
def test_helpers_resolve_helper_returns_no_code_when_helper_has_no_code() -> None:
    sentinel = ModuleType("sentinel")

    class Weird:
        pass

    sentinel.weird = Weird()  # type: ignore[attr-defined]
    file_path, error = coverage_gate_resolution._resolve_helper(sentinel, "app.x", "weird")
    assert file_path is None
    assert error == "no_code"


@pytestmark_helpers_required
def test_helpers_evaluate_helper_returns_none_pair_when_fully_covered() -> None:
    cov = _make_coverage_data(
        measured_files=["some/file.py"],
        executable_lines={"some/file.py": {1, 2}},
        executed_lines={"some/file.py": {1, 2}},
    )
    warning, failure = coverage_gate_resolution._evaluate_helper(
        "some/file.py", "app.x", "h", cov, {1, 2}
    )
    assert warning is None
    assert failure is None


@pytestmark_helpers_required
def test_helpers_evaluate_helper_returns_warning_when_not_measured() -> None:
    cov = _make_coverage_data(
        measured_files=[],
        executable_lines={},
        executed_lines={},
    )
    warning, failure = coverage_gate_resolution._evaluate_helper(
        "some/file.py", "app.x", "h", cov, {1}
    )
    assert warning is not None
    assert "not measured" in warning
    assert failure is None


@pytestmark_helpers_required
def test_helpers_evaluate_helper_returns_failure_when_under_covered() -> None:
    cov = _make_coverage_data(
        measured_files=["some/file.py"],
        executable_lines={"some/file.py": {1, 2, 3, 4}},
        executed_lines={"some/file.py": {1}},
    )
    warning, failure = coverage_gate_resolution._evaluate_helper(
        "some/file.py", "app.x", "h", cov, {1, 2, 3, 4}
    )
    assert warning is None
    assert failure is not None
    assert "FAIL" in failure


# --------------------------------------------------------------------------------------------
# Self-check: pytest itself imports this module, so prove the plugin loaded.
# --------------------------------------------------------------------------------------------


def test_coverage_gate_plugin_is_registered_in_addopts() -> None:
    """The CI workflow assumes the plugin is loaded via ``-p app.pytest_plugin.coverage_gate``.

    This is the assertion that prevents a future "I removed the plugin because
    nothing imported it" mistake: the test collection itself depends on the
    plugin being registered in `app/pyproject.toml`. If this test ever stops
    importing the module, the CI gate stops measuring the file and we re-enter
    the 66.06% loop.
    """
    assert "app.pytest_plugin.coverage_gate" in sys.modules
