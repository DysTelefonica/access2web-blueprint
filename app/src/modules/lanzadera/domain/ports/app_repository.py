"""AppRepositoryPort.

D7, D52, D85, DA-7. The single source of truth for the platform's
app catalog: identity, deployment topology (central vs office-nas),
office-presence requirement, and registration status. The
``list_active`` read path is the hot path for the delivery layer
(``GET /admin/apps`` and the ``list_visible_to`` permission check);
DA-8 caches it with a 5-minute TTL.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Protocol
from uuid import UUID

from app.src.modules.lanzadera.domain.app import App


class AppRepositoryPort(Protocol):
    """Persistence boundary for the app catalog."""

    async def get_by_id(self, app_id: int) -> App | None:
        """Return the app with ``app_id`` or ``None`` if it does not exist."""
        ...

    async def list_active(self) -> Sequence[App]:
        """Return every app with ``registration_status='active'``.

        Cached with TTL 5 min (DA-8); the cache is invalidated
        when an app's status changes (the ``list_visible_to`` call
        does not use the cache because the visibility check is
        per-user).
        """
        ...

    async def list_visible_to(self, user_id: UUID) -> Sequence[App]:
        """Return every app that ``user_id`` is allowed to see.

        Includes both active and pending apps the user has been
        explicitly granted access to via the assignment table.
        """
        ...