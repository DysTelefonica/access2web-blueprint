"""Tests for the DSN fixture's fail-not-skip contract (issue #882, PR-2).

Before issue #882, the integration DSN fixture used ``pytest.skip`` when the
Postgres URL was missing. The green CI rolled up without ever running the
integration tests — a false guarantee. PR-2 turned every such ``skip`` into a
``pytest.fail`` so a missing DSN shows up as a red build. These tests pin the
new contract: every code path in ``_integration_url_or_fail`` that used to
silently skip must now raise, and only the happy path returns the URL.

The fixture function lives at module level in ``test_user_repository_pg.py``
(deliberately, so the route ``tests/`` -> conftest.py -> fixture is minimal).
Importing it here is the smallest way to assert the contract without
duplicating the logic.
"""

from __future__ import annotations

import pytest

from tests.lanzadera.adapters.test_user_repository_pg import (
    INTEGRATION_URL_ENV,
    _integration_url_or_fail,
)


class _FailedError(Exception):
    """Marker exception asserted in place of ``pytest.fail.Exception``.

    The integration tests do not import ``pytest.fail.Exception`` directly
    because that attribute appeared under ``pytest._pytest.outcomes.Failed``
    before pytest 8 and as ``pytest.fail.Exception`` from pytest 8 onwards.
    The simpler cross-version assertion is ``pytest.raises(Exception)`` plus a
    message check; everything here is stdlib + the public pytest API.
    """


def test_dsn_fixture_fails_when_integration_flag_off(monkeypatch: pytest.MonkeyPatch) -> None:
    """Without ``APAP_INTEGRATION_ENABLED=1`` the fixture fails, not skips.

    This is the headline case the issue cites: a missing integration
    flag used to be a silent skip. After PR-2 it is a red build.
    """
    monkeypatch.delenv("APAP_INTEGRATION_ENABLED", raising=False)
    monkeypatch.delenv(INTEGRATION_URL_ENV, raising=False)
    with pytest.raises(pytest.fail.Exception, match="APAP_INTEGRATION_ENABLED is not 1"):
        _integration_url_or_fail()


def test_dsn_fixture_fails_when_flag_on_but_url_unset(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The flag is not enough — the URL itself must be set, or the fixture fails.

    Catches the other half of the previous skip behavior. Setting
    ``APAP_INTEGRATION_ENABLED=1`` without an ``APAP_INTEGRATION_URL``
    used to skip silently; now it fails with a message naming the
    missing variable so the operator does not have to read the script
    to find the gap.
    """
    monkeypatch.setenv("APAP_INTEGRATION_ENABLED", "1")
    monkeypatch.delenv(INTEGRATION_URL_ENV, raising=False)
    with pytest.raises(pytest.fail.Exception, match=f"{INTEGRATION_URL_ENV} is not set"):
        _integration_url_or_fail()


def test_dsn_fixture_returns_url_when_both_env_vars_set(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Happy path: the fixture returns the URL when both env vars are set.

    A test that only proves the failure paths is not enough — the
    fixture also has to keep working when the DSN is configured. This
    pins that contract: the same function that fails closed on missing
    config returns the URL unchanged when the config is present.
    """
    monkeypatch.setenv("APAP_INTEGRATION_ENABLED", "1")
    monkeypatch.setenv(INTEGRATION_URL_ENV, "postgresql://user:pw@host:5432/db")
    assert _integration_url_or_fail() == "postgresql://user:pw@host:5432/db"
