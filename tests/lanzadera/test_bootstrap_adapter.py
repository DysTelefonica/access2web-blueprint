"""Contract-conformance test for :class:`BootstrapAdminSource` + :class:`EnvAdminSourceAdapter`.

DA-6 + D91: the operator sets ``GLOBAL_ADMIN_EMAILS`` once at process
startup; the value is ``;``-separated, whitespace is trimmed, and an
unset variable returns an empty sequence (the caller treats that as a
no-op). The adapter does NOT raise on missing env — that is part of
the port contract (see :mod:`app.src.modules.lanzadera.ports.bootstrap_admin_source`).
"""

from __future__ import annotations

import pytest

from app.src.modules.lanzadera.adapters.bootstrap.env_admin_source_adapter import (
    EnvAdminSourceAdapter,
)
from app.src.modules.lanzadera.ports.bootstrap_admin_source import (
    BootstrapAdminSource,
)

ENV_KEY = "GLOBAL_ADMIN_EMAILS"


@pytest.fixture
def clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Remove GLOBAL_ADMIN_EMAILS from the environment for the duration of the test."""
    monkeypatch.delenv(ENV_KEY, raising=False)


def test_env_source_satisfies_protocol() -> None:
    """Static: the adapter implements the Protocol."""
    adapter: BootstrapAdminSource = EnvAdminSourceAdapter()
    assert hasattr(adapter, "list_initial_emails")
    assert callable(adapter.list_initial_emails)


def test_env_source_returns_empty_when_unset(clean_env: None) -> None:
    """DA-6: unset variable is a no-op, not an error."""
    adapter = EnvAdminSourceAdapter()
    assert adapter.list_initial_emails() == []


def test_env_source_returns_empty_when_empty_string(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(ENV_KEY, "")
    assert EnvAdminSourceAdapter().list_initial_emails() == []


def test_env_source_parses_single_email(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(ENV_KEY, "[email protected]")
    assert EnvAdminSourceAdapter().list_initial_emails() == ["[email protected]"]


def test_env_source_parses_multiple_emails(monkeypatch: pytest.MonkeyPatch) -> None:
    """DA-6: the value is ``;``-separated (not comma, not space)."""
    monkeypatch.setenv(ENV_KEY, "[email protected];[email protected];[email protected]")
    result = EnvAdminSourceAdapter().list_initial_emails()
    assert set(result) == {"[email protected]"}
    assert len(result) == 3


def test_env_source_strips_whitespace_around_addresses(monkeypatch: pytest.MonkeyPatch) -> None:
    """Trailing/leading whitespace around each address is trimmed."""
    monkeypatch.setenv(ENV_KEY, "  [email protected]  ;  [email protected]  ")
    result = EnvAdminSourceAdapter().list_initial_emails()
    assert set(result) == {"[email protected]"}


def test_env_source_skips_empty_segments(monkeypatch: pytest.MonkeyPatch) -> None:
    """``;;`` (double semicolon) is treated as a single separator, not an empty entry."""
    monkeypatch.setenv(ENV_KEY, "[email protected];;[email protected]")
    result = EnvAdminSourceAdapter().list_initial_emails()
    assert set(result) == {"[email protected]"}


def test_env_source_returns_empty_when_all_segments_blank(monkeypatch: pytest.MonkeyPatch) -> None:
    """All-whitespace segments collapse to empty."""
    monkeypatch.setenv(ENV_KEY, "  ;  ;  ")
    assert EnvAdminSourceAdapter().list_initial_emails() == []


def test_env_source_idempotent_across_calls(monkeypatch: pytest.MonkeyPatch) -> None:
    """DA-6: the source is pure — repeated calls return the same data."""
    monkeypatch.setenv(ENV_KEY, "[email protected];[email protected]")
    adapter = EnvAdminSourceAdapter()
    first = sorted(adapter.list_initial_emails())
    second = sorted(adapter.list_initial_emails())
    assert first == second
    assert first == ["[email protected]", "[email protected]"]
