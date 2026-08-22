"""ProfileRepositoryPort.

D22, D45, D46, DA-12, D110. The persistence boundary for the
``lanzadera.profiles`` table. The hot path is
``list_for_app(app_id)``, called by ``AssignmentRepository.effective_permissions``
on every permission check (DA-22). Cached by ``TtlCacheAdapter`` with
TTL 5 min (DA-8).

The seed migration 0003 (PR 3b) writes three default profiles per app:
``DEFAULT`` (the empty permissions catch-all), ``CALIDAD`` (the legacy
quality role), and ``TECNICO`` (the legacy technician role). The
``SinAcceso`` profile is enforced by the application layer as a
short-circuit (DA-12, H11).
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Protocol
from uuid import UUID

from app.src.modules.lanzadera.domain.profile import Profile


class ProfileRepositoryPort(Protocol):
    """Persistence boundary for user-app profiles."""

    async def list_for_app(self, app_id: int) -> Sequence[Profile]:
        """Return every active profile for ``app_id`` (DA-12, ordered by code)."""
        ...

    async def get_by_code(self, app_id: int, code: str) -> Profile | None:
        """Return the profile with ``(app_id, code)`` or ``None``."""
        ...

    async def create(self, profile: Profile) -> None:
        """Insert a new profile."""
        ...

    async def set_active(self, profile_id: UUID, active: bool) -> None:
        """Flip the ``active`` bit on a profile (D45)."""
        ...