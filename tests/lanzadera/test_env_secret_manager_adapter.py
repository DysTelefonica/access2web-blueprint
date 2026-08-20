"""Contract-conformance test for :class:`EnvSecretManagerAdapter`.

D25, D73, D11, CA-S2, CA-S4. The adapter reads the secret from
``os.environ`` and raises ``KeyError`` (no key in the message, per
CA-S4) when the key is missing. ``KeyError`` (not ``ValueError``) is
deliberate so the test-suite can ``except KeyError`` exact-match.

The protocol surface is ``SecretManagerPort.get(key: str) -> str``.
The previous round of the WU covered Protocol conformance via the
``FakeSecretManager`` inside the cipher tests; this file adds the
adapter-specific contract (env lookup, missing-key error, type check).
"""

from __future__ import annotations

import pytest

from app.src.modules.lanzadera.adapters.cross.secret_manager import (
    EnvSecretManagerAdapter,
)
from app.src.modules.lanzadera.domain.ports.secret_manager import SecretManagerPort

ENV_KEY = "USERS_DNI_CIPHER_TEST"  # test-only env var; never used in prod


@pytest.fixture
def clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(ENV_KEY, raising=False)


def test_env_secret_manager_satisfies_protocol() -> None:
    """Static: the adapter implements the Protocol."""
    adapter: SecretManagerPort = EnvSecretManagerAdapter()
    assert hasattr(adapter, "get")
    assert callable(adapter.get)


def test_env_secret_manager_returns_value_when_set(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(ENV_KEY, "the-actual-secret")
    adapter = EnvSecretManagerAdapter()
    assert adapter.get(ENV_KEY) == "the-actual-secret"


def test_env_secret_manager_raises_keyerror_when_unset(clean_env: None) -> None:
    """CA-S4: missing key is a hard error, not a default.

    The ``KeyError`` argument IS the key name (so the test-suite can
    route the 4xx), but it MUST NOT be echoed alongside any other
    secret value in the same log line.
    """
    adapter = EnvSecretManagerAdapter()
    with pytest.raises(KeyError) as exc_info:
        adapter.get(ENV_KEY)
    assert exc_info.value.args == (ENV_KEY,)


def test_env_secret_manager_returns_empty_string_when_set_empty(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """An empty value is a valid (if degenerate) value — the adapter returns ``""``.

    The application layer is responsible for rejecting empty secrets
    upstream (the ``NationalIdCipher`` already does so via Fernet
    validation). The adapter does not invent its own validation.
    """
    monkeypatch.setenv(ENV_KEY, "")
    adapter = EnvSecretManagerAdapter()
    assert adapter.get(ENV_KEY) == ""


def test_env_secret_manager_distinguishes_keys(monkeypatch: pytest.MonkeyPatch) -> None:
    """Different keys return different values; missing keys raise for the right key."""
    monkeypatch.setenv(ENV_KEY + "_A", "alpha")
    monkeypatch.setenv(ENV_KEY + "_B", "bravo")
    monkeypatch.delenv(ENV_KEY, raising=False)

    adapter = EnvSecretManagerAdapter()
    assert adapter.get(ENV_KEY + "_A") == "alpha"
    assert adapter.get(ENV_KEY + "_B") == "bravo"
    with pytest.raises(KeyError) as exc_info:
        adapter.get(ENV_KEY)
    assert exc_info.value.args == (ENV_KEY,)


def test_env_secret_manager_does_not_cache() -> None:
    """The adapter is pure — repeated calls hit ``os.environ`` each time."""
    import os

    adapter = EnvSecretManagerAdapter()
    os.environ[ENV_KEY] = "first"
    assert adapter.get(ENV_KEY) == "first"
    os.environ[ENV_KEY] = "second"
    assert adapter.get(ENV_KEY) == "second"
    del os.environ[ENV_KEY]
    with pytest.raises(KeyError):
        adapter.get(ENV_KEY)
