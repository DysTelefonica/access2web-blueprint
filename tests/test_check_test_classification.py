# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + testing-strategy epic #535
"""Tests for ``scripts/check_test_classification.py``.

Every rule has both a positive case (a layout that passes) and a negative case
(the same layout with the violation injected, asserting the gate reports it
with file and rule). If any negative case passes silently, the gate is
broken and the next layer-violation slips through.
"""

from __future__ import annotations

import contextlib
import io
import json
import sys
from pathlib import Path

import pytest

# Make ``scripts/`` importable without sys.path games.
SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import check_test_classification as gate  # noqa: E402  — sys.path tweak above

# --------------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------------


def _seed_layer(tmp_path: Path, layer: str, body: str) -> Path:
    target = tmp_path / "lanzadera" / layer
    target.mkdir(parents=True, exist_ok=True)
    p = target / "test_x.py"
    p.write_text(body, encoding="utf-8")
    return p


def _run(tmp_path: Path, *flags: str) -> tuple[int, str, str]:
    err, out = io.StringIO(), io.StringIO()
    with contextlib.redirect_stderr(err), contextlib.redirect_stdout(out):
        rc = gate.main(["--root", str(tmp_path / "lanzadera"), *flags])
    return rc, out.getvalue(), err.getvalue()


def _rule_hits(stderr: str, prefix: str) -> bool:
    return any(line.strip().startswith(prefix) for line in stderr.splitlines())


# --------------------------------------------------------------------------------------------
# Sample test bodies (one positive, one negative per rule)
# --------------------------------------------------------------------------------------------

# R1 — application/ with FakeFixtures vs. with MagicMock
APP_OK = (
    "from tests.lanzadera._fakes import FakeUserRepository\n"
    "def test_x() -> None:\n"
    "    assert FakeUserRepository() is not None\n"
)
APP_BAD = (
    "from unittest.mock import MagicMock\n"
    "def test_x() -> None:\n"
    "    assert MagicMock() is not None\n"
)
APP_STRING = 'def test_x() -> None:\n    """Do not use MagicMock here."""\n    assert True\n'

# R2 — delivery/ + admin_routes with/without auth gate
DELIVERY_OK_FIXTURE = (
    "from app.src.modules.lanzadera.delivery.http.admin_routes import register_routes\n"
    "import pytest\n"
    "@pytest.fixture\ndef auth_bypass(monkeypatch): ...\n"
    "def test_x() -> None:\n    assert True\n"
)
DELIVERY_OK_MARKER = (
    "from app.src.modules.lanzadera.delivery.http.admin_routes import register_routes\n"
    "import pytest\n"
    "@pytest.mark.requires_auth\ndef test_x() -> None:\n    assert True\n"
)
DELIVERY_BAD = (
    "from app.src.modules.lanzadera.delivery.http.admin_routes import register_routes\n"
    "def test_x() -> None:\n    assert True\n"
)
DELIVERY_NO_ADMIN = "def test_x() -> None:\n    assert True\n"


# --------------------------------------------------------------------------------------------
# R1 — HR-2: application/ tests must not import unittest.mock.MagicMock
# --------------------------------------------------------------------------------------------


class TestNoMagicMockInApplication:
    @pytest.mark.parametrize(
        "body, expected_violation",
        [(APP_OK, False), (APP_BAD, True)],
        ids=["fakes-ok", "magicmock-bad"],
    )
    def test_strict_mode(self, tmp_path: Path, body: str, expected_violation: bool) -> None:
        _seed_layer(tmp_path, "application", body)
        rc, _, err = _run(tmp_path, "--strict")
        assert (rc == 1) == expected_violation
        assert _rule_hits(err, "R1") == expected_violation

    def test_dry_run_prints_without_failing(self, tmp_path: Path) -> None:
        _seed_layer(tmp_path, "application", APP_BAD)
        rc, _, err = _run(tmp_path)  # default --dry-run
        assert rc == 0
        assert _rule_hits(err, "R1"), f"dry-run should still print; stderr:\n{err}"

    def test_string_mention_is_not_a_violation(self, tmp_path: Path) -> None:
        _seed_layer(tmp_path, "application", APP_STRING)
        rc, _, err = _run(tmp_path, "--strict")
        assert rc == 0, f"comment mention should pass; stderr:\n{err}"


# --------------------------------------------------------------------------------------------
# R2 — HR-6: delivery/ tests touching admin_routes declare auth_bypass or requires_auth
# --------------------------------------------------------------------------------------------


class TestAuthBypassInDelivery:
    @pytest.mark.parametrize(
        "body",
        [DELIVERY_OK_FIXTURE, DELIVERY_OK_MARKER, DELIVERY_NO_ADMIN],
        ids=["auth-bypass-fixture", "requires-auth-marker", "no-admin-routes"],
    )
    def test_passes(self, tmp_path: Path, body: str) -> None:
        _seed_layer(tmp_path, "delivery", body)
        rc, _, err = _run(tmp_path, "--strict")
        assert rc == 0, f"expected pass; got {rc}; stderr:\n{err}"

    def test_no_gate_flags_r2(self, tmp_path: Path) -> None:
        _seed_layer(tmp_path, "delivery", DELIVERY_BAD)
        rc, _, err = _run(tmp_path, "--strict")
        assert rc == 1
        assert _rule_hits(err, "R2"), f"expected R2 violation; stderr:\n{err}"


# --------------------------------------------------------------------------------------------
# R3 — HR-8: root-level tests must be in WIRE_ALLOWLIST
# --------------------------------------------------------------------------------------------


class TestRootPathLayer:
    def test_root_file_in_allowlist_passes(self, tmp_path: Path) -> None:
        (tmp_path / "lanzadera").mkdir()
        (tmp_path / "lanzadera" / next(iter(gate.WIRE_ALLOWLIST))).write_text(
            "def test_x() -> None:\n    assert True\n"
        )
        rc, _, err = _run(tmp_path, "--strict")
        assert rc == 0, f"allowlisted root file should pass; stderr:\n{err}"

    def test_root_file_not_in_allowlist_flags_r3(self, tmp_path: Path) -> None:
        (tmp_path / "lanzadera").mkdir()
        (tmp_path / "lanzadera" / "test_random_thing.py").write_text(
            "def test_x() -> None:\n    assert True\n"
        )
        rc, _, err = _run(tmp_path, "--strict")
        assert rc == 1
        assert _rule_hits(err, "R3"), f"expected R3 violation; stderr:\n{err}"

    def test_layer_subdir_file_always_passes(self, tmp_path: Path) -> None:
        _seed_layer(tmp_path, "domain", "def test_x() -> None:\n    assert True\n")
        rc, _, err = _run(tmp_path, "--strict")
        assert rc == 0, f"layer file should pass; stderr:\n{err}"


# --------------------------------------------------------------------------------------------
# CLI shape — argparse behavior
# --------------------------------------------------------------------------------------------


class TestCLI:
    def test_help_exits_zero(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            with pytest.raises(SystemExit) as exc:
                gate.main(["--help"])
        assert exc.value.code == 0

    def test_json_envelope_with_violation(self, tmp_path: Path) -> None:
        _seed_layer(tmp_path, "application", APP_BAD)
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            rc = gate.main(["--root", str(tmp_path / "lanzadera"), "--strict", "--json"])
        assert rc == 1
        envelope = json.loads(out.getvalue())
        assert envelope["gate"] == "test_classification"
        assert envelope["mode"] == "strict"
        assert envelope["status"] == "fail"
        assert envelope["indicators"]["violations"] == 1
        assert envelope["findings"][0]["rule"] == "R1"

    def test_json_envelope_clean(self, tmp_path: Path) -> None:
        (tmp_path / "lanzadera").mkdir()
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            rc = gate.main(["--root", str(tmp_path / "lanzadera"), "--json"])
        assert rc == 0
        envelope = json.loads(out.getvalue())
        assert envelope["mode"] == "dry-run"
        assert envelope["status"] == "pass"


# --------------------------------------------------------------------------------------------
# Real-repo smoke — the existing tests/lanzadera/ layout must remain a baseline
# --------------------------------------------------------------------------------------------


class TestRealRepoBaseline:
    """The current ``tests/lanzadera/`` layout must produce zero violations.

    Catches the gate itself drifting out of sync with the layout it guards.
    """

    def test_real_tests_lanzadera_is_clean(self) -> None:
        repo_root = Path(__file__).resolve().parent.parent
        tests_root = repo_root / "tests" / "lanzadera"
        if not tests_root.is_dir():
            pytest.skip("tests/lanzadera/ not present in this checkout")
        rc, _, err = _run(repo_root, "--strict")
        assert rc == 0, f"tests/lanzadera/ must remain clean; stderr:\n{err}"
