"""AssignmentRepositoryPort.

D22, D42, DA-12. The single source of truth for effective permissions
in the platform: a (user, app, profile) triple plus soft-deletion via
``revoked_at``. The application layer enforces ``SinAcceso``
exclusivity on top of this port — see DA-12 + H11.

The fourth method, ``effective_permissions``, is the read path the
delivery layer hits on every request and that is why the adapter is
required to cache it per ``(user_id, app_id)`` with TTL 60 s (DA-8).
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Protocol
from uuid import UUID

from app.src.modules.lanzadera.domain.assignment import Assignment


class AssignmentRepositoryPort(Protocol):
    """Persistence boundary for user-app-profile assignments."""

    async def create(self, user_id: UUID, app_id: int, profile_id: UUID) -> Assignment:
        """Persist a new assignment. Raises on duplicate or FK violation."""
        ...

    async def list_for_user(self, user_id: UUID) -> Sequence[Assignment]:
        """Return the live assignments for ``user_id`` (``revoked_at IS NULL``)."""
        ...

    async def list_for_app(self, app_id: int) -> Sequence[Assignment]:
        """Return the live assignments for ``app_id`` (``revoked_at IS NULL``)."""
        ...

    async def effective_permissions(self, user_id: UUID, app_id: int) -> Sequence[str]:
        """Return the capability names granted by the user's profile for ``app_id``.

        Result is cached per ``(user_id, app_id)`` with TTL 60 s
        (DA-8); the cache is invalidated on every mutation.
        """
        ...
