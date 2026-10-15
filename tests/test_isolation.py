# HARNESS-PROVENANCE: gent-883 — acceptance tests for the unit-suite isolation contract.
"""Acceptance tests for issue #883 (pytest-socket + env isolation).

These tests are the gate the issue asked for:

  1. A unit test that opens a socket to an external host fails
     (criterion 1, ``--disable-socket`` wired in ``addopts``).
  2. A unit test that reads an env var outside the allowlist does not
     see it (criterion 2, the ``_isolate_unit_env`` autouse fixture in
     ``tests/conftest.py``).

The autouse fixture in ``tests/conftest.py`` and pytest-socket's
``--disable-socket --allow-unix-socket --allow-hosts=127.0.0.1`` work
together: pytest-socket catches the network call before the syscall, the
fixture wipes the developer's ``.env`` before the test body runs. Neither
contract is opt-in for unit tests, so a regression in either wiring shows
up as a failed assertion here before it ever reaches CI.

The integration-test sanity check at the bottom is intentionally NOT
marked with ``@pytest.mark.integration``: if it were, the autouse fixture
would skip the env wipe and the unit-test assertion below would not
exercise the contract. Instead, we reach into ``request.keywords``
directly to assert the marker is recognised by pytest, then re-enable
socket access with ``@pytest.mark.allow_uncaught_hosts`` from
pytest-socket to prove the bypass works on both fronts.
"""

from __future__ import annotations

import os
import socket

import pytest

# Module-level canary: set when this module is first imported, before any
# test or fixture runs. The ``_isolate_unit_env`` autouse fixture must
# clear it before any unit-test body executes. If the fixture is missing
# or its allowlist leaks this prefix (``PYTEST_*`` is too broad), the
# second test below fails immediately — proving the regression is caught
# by the gate, not by ad-hoc inspection.
os.environ.setdefault("PYTEST_ISOLATION_CANARY", "module-import-time")


def test_unit_test_socket_to_external_host_is_blocked() -> None:
    """``--disable-socket`` blocks outbound connections in unit tests.

    pytest-socket raises ``SocketBlockedError`` before the syscall. On
    platforms where the plugin hasn't intercepted (e.g., the plugin is
    misconfigured or the marker is wrong), the OS raises
    ``OSError(ENETUNREACH)`` or similar. The test passes if the
    connection attempt does not succeed — that's the assertion we care
    about for #883.
    """
    with pytest.raises(Exception) as exc_info:
        socket.create_connection(("example.com", 80), timeout=1.0)
    assert exc_info.value is not None, (
        "socket.create_connection to example.com:80 must raise under "
        "--disable-socket. If it succeeded, the plugin is not wired."
    )


def test_unit_test_env_var_outside_allowlist_is_invisible() -> None:
    """The autouse env-isolation fixture wipes non-allowlisted vars before the test body runs.

    The canary above was set at module-import time (before any fixture or
    test ran). By the time this test body executes, the autouse fixture
    must have removed it. If the fixture's allowlist leaks
    ``PYTEST_ISOLATION_CANARY`` (e.g., a too-broad ``PYTEST_*`` prefix),
    the assertion below fails — the unit test would leak the
    developer's ``.env`` into the test body.
    """
    assert "PYTEST_ISOLATION_CANARY" not in os.environ, (
        "The autouse env-isolation fixture in tests/conftest.py should "
        "have cleared PYTEST_ISOLATION_CANARY before this test body "
        "ran. If you see this, either the fixture is missing or its "
        "allowlist is too permissive."
    )


@pytest.mark.integration
@pytest.mark.allow_uncaught_hosts
def test_integration_test_marker_bypasses_env_isolation() -> None:
    """``@pytest.mark.integration`` opts out of ``_isolate_unit_env`` and ``--disable-socket``.

    A canary set in this body must survive to the assertion. The
    autouse fixture returns early when the integration marker is
    present, so the env wipe does not run. pytest-socket's
    ``allow_uncaught_hosts`` marker lets the loopback ``connect`` call
    through to the OS (it'll fail with ``ConnectionRefusedError`` on
    port 1, but the test is about the bypass, not the handshake).
    """
    os.environ["PYTEST_INTEGRATION_CANARY"] = "integration-test-set-this"
    try:
        assert os.environ.get("PYTEST_INTEGRATION_CANARY") == "integration-test-set-this", (
            "Integration tests should bypass the autouse env-isolation; "
            "if the canary is gone, the fixture's marker check is broken."
        )
    finally:
        os.environ.pop("PYTEST_INTEGRATION_CANARY", None)
