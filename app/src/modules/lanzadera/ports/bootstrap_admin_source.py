"""BootstrapAdminSource — D21, D42, D48, D91, DA-6.

Driven port that lists the e-mail addresses of the platform's first
global admins. The composition root invokes this at process startup;
the result feeds the bootstrap seed path (D21 + D42 + D48) that
creates the initial ``users`` and ``global_admins`` rows. Idempotent:
if the variable is unset, the source returns an empty sequence and
the bootstrap is a no-op.

The port is intentionally tiny: it returns addresses, not users.
The seed-mutation lives in the application layer (composition root
+ the seed step of the Alembic 0001 migration) so the source stays
substitutable (Vault, secret manager, LDAP) without touching domain.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Protocol


class BootstrapAdminSource(Protocol):
    """Driven port for the initial global-admin e-mail list."""

    def list_initial_emails(self) -> Sequence[str]:
        """Return every e-mail address that should bootstrap a global admin.

        An empty sequence means the operator has not configured the
        bootstrap yet — the caller MUST treat that as a no-op and not
        raise. The list is unordered; the caller applies its own
        ordering and dedup policy.
        """
        ...
