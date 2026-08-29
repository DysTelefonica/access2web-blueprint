"""``EnvSecretManagerAdapter`` — D25, D73, D11, CA-S2, CA-S4.

Reads the secret material from ``os.environ``. The MVP source for
``SecretManagerPort``. Production swaps in a secret-manager-backed
adapter (Vault / AWS Secrets Manager) without touching the domain.

The adapter does NOT cache: the secret manager itself is the
authoritative source and the composition root is the only caller.
The key never appears in logs, exception messages, or CLI arguments
(CA-S4): a missing key raises ``KeyError`` (no key in the message),
and a non-str value propagates the type as ``TypeError``.

``KeyError`` (not ``ValueError``) is deliberate: the port contract
caller maps the missing-key to a 4xx/5xx response, and a ``KeyError``
makes the test-suite grep ``except KeyError`` exact-match the contract.
"""

from __future__ import annotations

import base64
import hashlib
import os

from cryptography.fernet import Fernet

from app.src.modules.lanzadera.domain.ports.secret_manager import SecretManagerPort


class EnvSecretManagerAdapter(SecretManagerPort):
    """Driven adapter: read the secret from ``os.environ``."""

    def get(self, key: str) -> str:
        """Return the raw secret value bound to ``key``.

        Raises:
            KeyError: when ``key`` is not configured. The key name is
                the exception argument so callers can route the 4xx,
                but it MUST NOT be echoed alongside any other secret
                value in the same log line (CA-S4).
        """
        try:
            value = os.environ[key]
        except KeyError:
            raise KeyError(key) from None
        if not isinstance(value, str):
            # ``os.environ`` always returns ``str`` on POSIX, but the
            # contract documents the precondition — non-str would be a
            # deployment bug worth surfacing loudly.
            raise TypeError(f"secret {key!r} is not a string; check the deployment env") from None
        return value

    def _fernet(self) -> Fernet:
        """Lazily-build the Fernet cipher from the configured key."""
        if not hasattr(self, "_fernet_cache"):
            raw = os.environ.get("FERNET_KEY")
            if raw:
                key: bytes = base64.urlsafe_b64decode(raw.encode())
            else:
                secret = os.environ.get("SECRET_KEY", "dev-only-fallback")
                key = hashlib.sha256(secret.encode()).digest()
            self._fernet_cache = Fernet(base64.urlsafe_b64encode(key))
        return self._fernet_cache  # type: ignore[has-type]

    def encrypt(self, plaintext: str) -> bytes:
        """Symmetric encrypt plaintext via Fernet.

        Used by the create_user use case to encrypt national_id before
        storage. Key source: FERNET_KEY env var (base64-encoded 32 bytes) or
        SECRET_KEY (SHA-256 derived). Raises KeyError if FERNET_KEY is set
        but unreadable.
        """
        return self._fernet().encrypt(plaintext.encode("utf-8"))


__all__ = ["EnvSecretManagerAdapter"]
