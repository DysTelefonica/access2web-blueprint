"""PostgreSQL adapter for :class:`AssignmentRepositoryPort`.

DA-12 + D22 + D42 + H11: assignments are the only source of truth for
effective permissions; revocation is a soft-delete (``revoked_at``).

**Skeleton** — the bodies raise ``NotImplementedError`` until migrations
``0001_core_schema`` (already on main) + ``0003_seed_assignments`` (next
WU) provide the ``user_app_assignments`` table. The contract is
declared today so the application layer can pin its tests against
``typing.Protocol`` structural conformance.
"""

from __future__ import annotations

from collections.abc import Sequence
from uuid import UUID

from app.src.modules.lanzadera.domain.assignment import Assignment
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)


class AssignmentRepositoryPg(AssignmentRepositoryPort):
    """SQLAlchemy Core + asyncpg implementation. Skeleton — see module docstring."""

    async def create(self, user_id: UUID, app_id: int, profile_id: UUID) -> Assignment:
        raise NotImplementedError(
            "AssignmentRepositoryPg.create is awaiting migration 0003_seed_assignments"
        )

    async def list_for_user(self, user_id: UUID) -> Sequence[Assignment]:
        raise NotImplementedError(
            "AssignmentRepositoryPg.list_for_user is awaiting migration 0003_seed_assignments"
        )

    async def list_for_app(self, app_id: int) -> Sequence[Assignment]:
        raise NotImplementedError(
            "AssignmentRepositoryPg.list_for_app is awaiting migration 0003_seed_assignments"
        )

    async def effective_permissions(self, user_id: UUID, app_id: int) -> Sequence[str]:
        raise NotImplementedError(
            "AssignmentRepositoryPg.effective_permissions awaits 0003 + TtlCache adapter"
        )
