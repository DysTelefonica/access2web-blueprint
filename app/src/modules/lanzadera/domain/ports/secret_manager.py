"""Cross-cutting port for secret retrieval (CA-S2, D25).

The MVP adapter reads from process environment (``EnvSecretManagerAdapter``);
production swaps in Vault / AWS Secrets Manager without touching the domain
(D73, D11). The secret *value* never appears in logs or CLI arguments
(CA-S4): consumers receive opaque bytes via :meth:`get` and pass them to
KDFs/ciphers that hold them in-memory only.
"""

from __future__ import annotations

from typing import Protocol


class SecretManagerPort(Protocol):
    """Return the raw secret bytes for ``key`` or raise."""

    def get(self, key: str) -> str:
        """Return the secret material bound to ``key``.

        Raises:
            KeyError: when ``key`` is not configured in the underlying
                backend (e.g. env-var missing in dev). Callers MUST
                translate this to a user-facing 4xx/5xx, never log the
                key name alongside any other secret.
        """
        ...


SecretManager = SecretManagerPort
