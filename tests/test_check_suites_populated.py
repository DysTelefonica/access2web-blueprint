"""Tests for `scripts/check_suites_populated.py` (issue #882, PR-1).

The guard runs `pytest --collect-only -q -m <expr>` for each of the three
mandatory suites (`unit`, `integration`, `e2e_smoke`) and fails closed if any
suite collects zero tests. These tests prove three things about the contract:

  1. A clean tree with all three suites populated passes.
  2. Renaming a marker (or removing all tests from a suite) makes the
     guard fail with an actionable message — directly satisfying the
     acceptance criterion "Renombrar el marker `integration` hace fallar
     el guard".
  3. pytest itself failing during collection (unregistered marker,
     unparseable file) is reported rather than silently passing.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import check_suites_populated as guard  # noqa: E402  — sys.path tweak above


def _load_guard_in_tmp_repo(tmp_path: Path) -> tuple[object, Path]:
    """Materialise a minimal pytest project under ``tmp_path`` and import
    the guard against it.

    The copy carries just enough infrastructure for ``pytest -c
    app/pyproject.toml --rootdir app --collect-only -q`` to succeed: the
    pyproject with the three markers registered, the tests directory
    layout pytest looks for, and the user-supplied test file.
    """
    repo = tmp_path
    (repo / "app").mkdir()
    (repo / "tests").mkdir()
    pyproject = repo / "app" / "pyproject.toml"
    pyproject.write_text(
        "[pytest]\n"
        'testpaths = ["../tests"]\n'
        "markers = [\n"
        '  "integration: touches a real database or another out-of-process dependency",\n'
        '  "e2e_smoke: gate-smoke tests that run real scripts as subprocesses",\n'
        "]\n"
        'python_files = ["test_*.py"]\n'
        'python_classes = ["Test*"]\n'
        'python_functions = ["test_*"]\n'
        "strict_markers = true\n"
    )
    return guard, repo


def _write_test(repo: Path, marker: str | None, body: str = "    pass\n") -> None:
    """Write one synthetic test, optionally carrying the given marker."""
    tests_dir = repo / "tests"
    tests_dir.mkdir(exist_ok=True)
    decorator = f"@pytest.mark.{marker}\n" if marker else ""
    (tests_dir / "test_synthetic.py").write_text(
        f"import pytest\n\n{decorator}def test_one():\n{body}"
    )


def test_check_suites_returns_empty_when_every_suite_is_populated(tmp_path: Path) -> None:
    """A clean repo with all three suites populated produces zero violations."""
    _, repo = _load_guard_in_tmp_repo(tmp_path)
    # unit: an unmarked test
    _write_test(repo, marker=None)
    # integration: one @pytest.mark.integration test
    (repo / "tests" / "test_int.py").write_text(
        "import pytest\n@pytest.mark.integration\ndef test_int_one():\n    pass\n"
    )
    # e2e_smoke: one @pytest.mark.e2e_smoke test
    (repo / "tests" / "test_smoke.py").write_text(
        "import pytest\n@pytest.mark.e2e_smoke\ndef test_smoke_one():\n    pass\n"
    )
    violations = guard.check_suites(repo)
    assert violations == [], f"expected no violations for a fully-populated repo, got: {violations}"


@pytest.mark.parametrize(
    "missing_marker",
    ["integration", "e2e_smoke"],
)
def test_check_suites_fails_when_a_marker_suite_is_empty(
    tmp_path: Path, missing_marker: str
) -> None:
    """Renaming or removing the marker for one suite makes the guard fail.

    Directly satisfies the acceptance criterion "Renombrar el marker
    `integration` hace fallar el guard." We simulate "renamed" by not
    including any test with that marker.
    """
    _, repo = _load_guard_in_tmp_repo(tmp_path)
    # unit: one unmarked test (keeps the unit suite populated)
    _write_test(repo, marker=None)
    # integration: one marked test (unless THIS is the missing one)
    if missing_marker != "integration":
        (repo / "tests" / "test_int.py").write_text(
            "import pytest\n@pytest.mark.integration\ndef test_int_one():\n    pass\n"
        )
    # e2e_smoke: one marked test (unless THIS is the missing one)
    if missing_marker != "e2e_smoke":
        (repo / "tests" / "test_smoke.py").write_text(
            "import pytest\n@pytest.mark.e2e_smoke\ndef test_smoke_one():\n    pass\n"
        )
    violations = guard.check_suites(repo)
    assert any(missing_marker in v for v in violations), (
        f"expected a violation naming the empty suite {missing_marker!r}, got: {violations}"
    )
    assert all(missing_marker in v for v in violations if missing_marker in v), (
        f"each violation for {missing_marker!r} must name the suite so the "
        f"operator can act on it. Got: {violations}"
    )


def test_check_suites_fails_when_unit_suite_is_empty(tmp_path: Path) -> None:
    """If every test carries a non-unit marker, the unit suite is empty.

    Hard to construct on a real repo (everything is unit by exclusion),
    so we synthesise it: two tests, both marked `integration`, no unmarked
    tests. The guard should report `unit` as empty.
    """
    _, repo = _load_guard_in_tmp_repo(tmp_path)
    (repo / "tests" / "test_int1.py").write_text(
        "import pytest\n@pytest.mark.integration\ndef test_one():\n    pass\n"
    )
    (repo / "tests" / "test_int2.py").write_text(
        "import pytest\n@pytest.mark.integration\ndef test_one():\n    pass\n"
    )
    violations = guard.check_suites(repo)
    assert any(v.startswith("unit:") for v in violations), (
        f"expected a `unit:` violation when every test is non-unit, got: {violations}"
    )


def test_check_suites_reports_pytest_collection_failure(tmp_path: Path) -> None:
    """An unregistered marker under --strict-markers fails pytest at COLLECTION.

    The guard must report this as a violation (a non-zero collection
    returncode), not as zero tests. Otherwise a typo in a marker name
    would look like an empty suite and the guard would mask a real
    infrastructure failure.
    """
    _, repo = _load_guard_in_tmp_repo(tmp_path)
    _write_test(repo, marker=None)
    (repo / "tests" / "test_unreg.py").write_text(
        "import pytest\n@pytest.mark.not_registered_in_pyproject_toml\ndef test_one():\n    pass\n"
    )
    violations = guard.check_suites(repo)
    assert violations, "expected at least one violation when pytest fails at COLLECTION"
    assert any("pytest collection failed" in v for v in violations), (
        f"expected the collection-failed message, got: {violations}"
    )


def test_check_suites_violations_name_the_suite_and_say_what_to_do(tmp_path: Path) -> None:
    """Every violation message names the empty suite AND tells the operator what to do.

    A guard that just reports "0 tests" forces the operator to read the
    script to find which suite failed and what to fix. Naming the suite
    and the fix path makes the failure actionable in the CI log.
    """
    _, repo = _load_guard_in_tmp_repo(tmp_path)
    # Empty repo: no tests at all, so every suite collects 0.
    violations = guard.check_suites(repo)
    assert violations, "an empty repo must produce violations"
    # Each suite name must appear in at least one violation.
    for suite_name in ("unit", "integration", "e2e_smoke"):
        assert any(suite_name in v for v in violations), (
            f"violations must name each empty suite ({suite_name!r}); got: {violations}"
        )
        # Each violation must include a directive ("Add a test…", "edit
        # SUITES…", or the collection-failure hint).
        matching = [v for v in violations if suite_name in v]
        assert any(
            "Add a test" in v or "edit SUITES" in v or "pytest collection" in v for v in matching
        ), f"violations for {suite_name} must tell the operator what to do: {matching}"


def test_main_exits_zero_when_every_suite_is_populated(tmp_path: Path) -> None:
    """`main()` returns 0 when all three suites collected at least one test."""
    _, repo = _load_guard_in_tmp_repo(tmp_path)
    _write_test(repo, marker=None)
    (repo / "tests" / "test_int.py").write_text(
        "import pytest\n@pytest.mark.integration\ndef test_one():\n    pass\n"
    )
    (repo / "tests" / "test_smoke.py").write_text(
        "import pytest\n@pytest.mark.e2e_smoke\ndef test_one():\n    pass\n"
    )
    assert guard.main([str(repo)]) == 0


def test_main_exits_nonzero_when_a_suite_is_empty(tmp_path: Path) -> None:
    """`main()` returns 1 with at least one empty suite, exactly the
    failure the guard exists to catch."""
    _, repo = _load_guard_in_tmp_repo(tmp_path)
    # Only unit tests, no integration or e2e_smoke.
    _write_test(repo, marker=None)
    assert guard.main([str(repo)]) == 1
