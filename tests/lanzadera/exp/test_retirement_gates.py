"""CAP-057 retirement gate — Win32 network.

Issue #264 (WU L01) ships the first retirement gate for the
expedientes-web-migration plan. The gate prevents the Win32 network
path (\\\\server\\\\share\\\\file.accdb, the access backend's file
share) from re-appearing in ``app/src/`` once the retirement lands.

The spec (``openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md``)
says: "the access backend's network path... is the dependency that
physically prevents the platform from running on Linux". The retirement
replaces ``win32file`` / ``win32net`` with ``DocumentStoragePort``
(CAP-057). The gate is the test that prevents reintroduction.

This test mirrors the existing ``tests/lanzadera/test_crap_retired_guard.py``
(issue #266, DG-12) and ``scripts/check_legacy_hashes.py`` (DA-13).
The pattern is the same: a script that walks the AST and reports
forbidden symbols, plus a test that calls the script and asserts the
surface is clean. One gate per WU until the other 6 CAPs (058-063)
land as ``app/src/`` grows the corresponding legacy.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
SCRIPT = REPO_ROOT / "scripts" / "check_legacy_retirement.py"


def test_cap_057_script_imports_and_prints_help() -> None:
    """The runner is importable and ``--help`` works (sanity check)."""
    result = subprocess.run(
        [sys.executable, str(SCRIPT), "--help"],
        capture_output=True,
        text=True,
        check=False,
        cwd=REPO_ROOT,
    )
    assert result.returncode == 0
    assert "CAP-057" in result.stdout or "retirement" in result.stdout.lower()


def test_cap_057_walk_target_does_not_exist_returns_zero() -> None:
    """A walk over a non-existent root returns zero violations (graceful)."""
    result = subprocess.run(
        [sys.executable, str(SCRIPT), "--root", "/nonexistent/path", "--gate", "CAP-057"],
        capture_output=True,
        text=True,
        check=False,
        cwd=REPO_ROOT,
    )
    # Script doesn't fail on missing root — ``app/src/`` is just not
    # walked. Exit 0 because there are zero violations (the contract is
    # "fail only on a violation found", not "fail on a missing path").
    assert result.returncode == 0


def test_cap_057_clean_app_src_passes() -> None:
    """Walking ``app/src/`` against the CAP-057 forbidden set must produce zero violations.

    This is the live pin: any reintroduction of ``win32file``,
    ``win32net``, ``win32com``, ``MSComctlLib``, ``WSAStartup``, etc. to
    ``app/src/`` flips this test red. The legacy code lives in the
    access backend; the platform's Python source MUST stay clear.
    """
    result = subprocess.run(
        [sys.executable, str(SCRIPT), "--root", str(REPO_ROOT), "--gate", "CAP-057"],
        capture_output=True,
        text=True,
        check=False,
        cwd=REPO_ROOT,
    )
    assert result.returncode == 0, (
        f"CAP-057 gate failed:\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
    )
    assert "OK" in result.stdout


def test_cap_057_forbidden_symbols_are_a_frozenset() -> None:
    """The script exposes the forbidden set as a frozenset so the test can introspect it.

    A test that imports the module directly (not via subprocess) and
    asserts the set is non-empty + has the key markers. This catches
    the case where someone edits the set to be empty (gate becomes
    no-op) without realizing it.
    """
    import importlib.util

    spec = importlib.util.spec_from_file_location("check_legacy_retirement", SCRIPT)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)  # type: ignore[union-attr]

    assert isinstance(module.FORBIDDEN_SYMBOLS_CAP057, frozenset)
    assert len(module.FORBIDDEN_SYMBOLS_CAP057) > 0
    # The markers that the issue body names explicitly:
    assert "win32file" in module.FORBIDDEN_SYMBOLS_CAP057
    assert "win32net" in module.FORBIDDEN_SYMBOLS_CAP057
    assert "win32com" in module.FORBIDDEN_SYMBOLS_CAP057
    assert "MSComctlLib" in module.FORBIDDEN_SYMBOLS_CAP057  # CAP-059 overlaps


def test_cap_057_violation_is_reported(tmp_path: Path) -> None:
    """A file containing a forbidden symbol MUST be reported (script's AST walker works).

    Creates a throwaway file under ``app/src/`` (the runner walks it),
    runs the script, asserts the return code is 1 and the file/symbol
    appear in stderr. The file is removed after the test so the rest
    of the suite is unaffected.
    """
    target = REPO_ROOT / "app" / "src" / "_cap057_smoke_probe.py"
    target.write_text("def _smoke() -> None:\n    import win32file  # CAP-057 forbidden\n")
    try:
        result = subprocess.run(
            [sys.executable, str(SCRIPT), "--root", str(REPO_ROOT), "--gate", "CAP-057"],
            capture_output=True,
            text=True,
            check=False,
            cwd=REPO_ROOT,
        )
        assert result.returncode == 1, (
            f"expected gate to fail; got rc={result.returncode}\n"
            f"STDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
        )
        assert "_cap057_smoke_probe.py" in result.stdout
        assert "win32file" in result.stdout
    finally:
        target.unlink()
