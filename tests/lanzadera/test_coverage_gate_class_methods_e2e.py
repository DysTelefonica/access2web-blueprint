# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + DA-13 — class-method coverage e2e
"""Live pytest-cov evaluation of CRITICAL_TARGETS whose owner is a class.

W24 (PR #453) updated the plugin's helper names to
``CredentialHasherArgon2id.hash`` and ``.verify`` (class methods, not free
functions), and W28 (#458) added ``_resolve_helper`` so the orchestrator
accepts an owner name and walks ``getattr(owner, helper_name)``. The
existing live-subprocess test
(``test_live_pytest_cov_fails_for_undercovered_exact_target``) only
exercises owner=None — a regression in the class-method resolution path
would pass it and only break in CI.

This module fills that gap: a synthetic class with two methods, an
under-covered ``verify`` branch, and a real ``pytest -p
app.pytest_plugin.coverage_gate --cov=<module>`` subprocess run that
asserts ``session.exitstatus == 1`` (Hard Rule 8 contract).

The synthetic module lives in ``tmp_path`` so the live subprocess sees a
clean module-to-path map and the gate's ``_resolve_helper`` resolves the
target the same way it resolves ``CredentialHasherArgon2id`` in
production. No fixture or import surface is shared with the real
adapter: the test only cares about the orchestrator's owner-aware code
path.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


def test_live_pytest_cov_fails_for_undercovered_class_method(tmp_path: Path) -> None:
    """Owner = class name; one method 100% covered, the other unmeasured.

    The plugin must flip ``session.exitstatus`` to 1 when the
    class-method ``verify`` is unmeasured and the orchestrator walks the
    ``_resolve_helper`` branch with a non-None owner.
    """
    (tmp_path / "hasher_module.py").write_text(
        "class Hasher:\n"
        "    def __init__(self) -> None:\n"
        "        self._salt = 'static'\n"
        "    def hash(self, password: str) -> str:\n"
        "        if not password:\n"
        "            raise ValueError('empty')\n"
        "        return 'hashed:' + password\n"
        "    def verify(self, password: str, password_hash: str) -> bool:\n"
        "        if not password or not password_hash:\n"
        "            return False\n"
        "        return password_hash == 'hashed:' + password\n",
        encoding="utf-8",
    )
    (tmp_path / "test_hasher.py").write_text(
        "from hasher_module import Hasher\n"
        "\n"
        "def test_hash() -> None:\n"
        "    h = Hasher()\n"
        "    assert h.hash('foo') == 'hashed:foo'\n",
        encoding="utf-8",
    )
    (tmp_path / "conftest.py").write_text(
        "import app.pytest_plugin.coverage_gate as gate\n"
        "gate.CRITICAL_TARGETS = (\n"
        "    ('hasher_module', 'Hasher', 'hash'),\n"
        "    ('hasher_module', 'Hasher', 'verify'),\n"
        ")\n",
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
            "--cov=hasher_module",
            "test_hasher.py",
        ],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert (result.returncode, "coverage_gate FAIL" in result.stderr) == (1, True)


def test_live_pytest_cov_silent_for_fully_covered_class_method(tmp_path: Path) -> None:
    """Owner = class name; both methods exercised — gate must stay silent.

    Counter-test for the FAIL path above: when the same class has both
    methods exercised (including their empty-input branches), the
    orchestrator's owner-aware resolution must return without emitting
    FAIL or mutating ``session.exitstatus``.
    """
    (tmp_path / "hasher_module.py").write_text(
        "class Hasher:\n"
        "    def __init__(self) -> None:\n"
        "        self._salt = 'static'\n"
        "    def hash(self, password: str) -> str:\n"
        "        if not password:\n"
        "            raise ValueError('empty')\n"
        "        return 'hashed:' + password\n"
        "    def verify(self, password: str, password_hash: str) -> bool:\n"
        "        if not password or not password_hash:\n"
        "            return False\n"
        "        return password_hash == 'hashed:' + password\n",
        encoding="utf-8",
    )
    (tmp_path / "test_hasher.py").write_text(
        "import pytest\n"
        "from hasher_module import Hasher\n"
        "\n"
        "def test_hash_happy_path() -> None:\n"
        "    h = Hasher()\n"
        "    assert h.hash('foo') == 'hashed:foo'\n"
        "\n"
        "def test_hash_rejects_empty() -> None:\n"
        "    h = Hasher()\n"
        "    with pytest.raises(ValueError):\n"
        "        h.hash('')\n"
        "\n"
        "def test_verify_happy_path() -> None:\n"
        "    h = Hasher()\n"
        "    digest = h.hash('foo')\n"
        "    assert h.verify('foo', digest) is True\n"
        "\n"
        "def test_verify_rejects_empty() -> None:\n"
        "    h = Hasher()\n"
        "    assert h.verify('', 'something') is False\n"
        "    assert h.verify('foo', '') is False\n",
        encoding="utf-8",
    )
    (tmp_path / "conftest.py").write_text(
        "import app.pytest_plugin.coverage_gate as gate\n"
        "gate.CRITICAL_TARGETS = (\n"
        "    ('hasher_module', 'Hasher', 'hash'),\n"
        "    ('hasher_module', 'Hasher', 'verify'),\n"
        ")\n",
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
            "--cov=hasher_module",
            "test_hasher.py",
        ],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert "coverage_gate FAIL" not in result.stderr
    assert result.returncode == 0
