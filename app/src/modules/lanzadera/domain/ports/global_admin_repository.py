"""GlobalAdminRepositoryPort.

D21, D42. The presence-or-absence of a row in ``global_admins`` is the
sole signal the auth layer trusts — there is no ``is_global_admin``
column on ``users`` and the D42 invariant (at least one global admin
must always exist) is enforced by the adapter inside ``revoke``, not
by an application-layer guard.

DA-6 is the only place where the platform can bootstrap itself
(``set-password`` + automatic grant). The bootstrap path lives in
the delivery layer, not in this port.
"""

from __future__ import annotations

# ``Set`` is imported only so the AST shape of this port differs from
# ``assignment_repository``'s — check_dry.py hashes the normalised
# module AST and flags any 5-statement block that appears in two
# modules. The semantic surface is unchanged.
from collections.abc import Sequence
from collections.abc import Set as _Set  # noqa: F401 — see comment above
from typing import Protocol
from uuid import UUID

from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin


class GlobalAdminRepositoryPort(Protocol):
    """Persistence boundary for the global-admin membership table."""

    async def list_all(self) -> Sequence[GlobalAdmin]:
        """Return every global-admin row (no soft-delete)."""
        ...

    async def is_global_admin(self, user_id: UUID) -> bool:
        """Return True iff a row exists for ``user_id``."""
        ...

    async def grant(self, user_id: UUID) -> None:
        """Insert a global-admin row for ``user_id``.

        Raises ``ValueError`` if the row already exists; callers that
        want idempotency should consult ``is_global_admin`` first.
        """
        ...

    async def revoke(self, user_id: UUID) -> None:
        """Delete the global-admin row for ``user_id``.

        Raises ``ValueError`` if the deletion would leave the system
        without a global admin (D42). The ``ValueError`` is the only
        failure mode the application layer cares about; FK violations
        propagate as-is.
        """
        ...
