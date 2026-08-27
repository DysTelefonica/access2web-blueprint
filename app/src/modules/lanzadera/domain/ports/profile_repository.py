# DA-12, D22, D45 — ProfileRepositoryPort.
"""Profile persistence boundary (DA-12, D22, D45, D46).

Profiles are scoped under an app: a profile belongs to exactly one
app and is identified by ``(app_id, code)``. The Postgres adapter
lands in W01 (#44); the in-memory fake lives in
``tests/lanzadera/auth/_fakes.py`` for the reset-flow and in
``tests/lanzadera/delivery/_fakes.py`` for the admin router.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Protocol
from uuid import UUID

from app.src.modules.lanzadera.domain.profile import Profile


class ProfileRepositoryPort(Protocol):
    """Persistence boundary for app-scoped profiles."""

    async def list_for_app(self, app_id: int) -> Sequence[Profile]:
        """Return every profile of ``app_id`` ordered by ``code``."""
        ...

    async def get_by_code(self, app_id: int, code: str) -> Profile | None:
        """Return the profile identified by ``(app_id, code)`` or ``None``."""

    async def get_by_id(self, profile_id: UUID) -> Profile | None:
        """Return the profile with ``id == profile_id`` or ``None``."""
        ...

    async def create(self, profile: Profile) -> None:
        """Persist a new ``Profile``."""
        ...

    async def set_active(self, app_id: int, code: str, *, active: bool) -> None:
        """Toggle ``active`` for the profile ``(app_id, code)``.

        Used by the admin router's ``POST /admin/apps/{app_id}/profiles/{code}/deactivate``
        and its reverse. Touches ``updated_at`` on the underlying row.
        """
        ...
