"""``EnvAdminSourceAdapter`` — DA-6, D91.

Reads ``os.environ["GLOBAL_ADMIN_EMAILS"]`` at process startup. The
value is split on ``;`` (DA-6) and stripped of whitespace; empty
inputs and an unset variable both return ``[]`` so the caller treats
the bootstrap as a no-op (per DA-6 + D48).

The adapter is intentionally environment-only for v1. Production
swaps in a secret-manager-backed source without touching the domain
(D73, D11) — the port contract does not change.
"""

from __future__ import annotations

import os

from app.src.modules.lanzadera.ports.bootstrap_admin_source import (
    BootstrapAdminSource,
)

_ENV_KEY = "GLOBAL_ADMIN_EMAILS"
_SEPARATOR = ";"


class EnvAdminSourceAdapter(BootstrapAdminSource):
    """Driven adapter: read the bootstrap e-mail list from the environment."""

    def list_initial_emails(self) -> list[str]:
        """Return the ``;``-separated addresses from ``$GLOBAL_ADMIN_EMAILS``.

        Empty if the variable is unset or contains only empty entries.
        Whitespace around each address is stripped.
        """
        raw = os.environ.get(_ENV_KEY, "")
        if not raw:
            return []
        return [part.strip() for part in raw.split(_SEPARATOR) if part.strip()]
