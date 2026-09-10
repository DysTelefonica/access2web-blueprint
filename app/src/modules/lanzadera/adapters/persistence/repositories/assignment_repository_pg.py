# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp W02 (#45)
# DA-12, D22, D42, H11 — Postgres adapter for the assignment port.
"""Async Postgres adapter for the ``AssignmentRepositoryPort`` Protocol.

Implements the contract declared in
``app.src.modules.lanzadera.domain.ports.assignment_repository.AssignmentRepositoryPort``
(DA-12, D22, D42, H11). Assignments are the only source of truth for
effective permissions in the platform; revocation is a soft-delete
(``revoked_at`` is filled when a global admin revokes the profile;
the row stays in the table so audit reads see the full lifecycle).

The hot path ``effective_permissions`` is the read the delivery
layer hits on every request (DA-8); a 60-second TTL cache lives at a
higher layer (the application side), not in this adapter. The
adapter is responsible for the SQL semantics; the cache sits between
this object and the FastAPI dependency.

The four-method contract:
- ``create`` — INSERT a fresh row.
- ``list_for_user`` — SELECT all live rows for ``user_id``
  (``revoked_at IS NULL``).
- ``list_for_app`` — SELECT all live rows for ``app_id``
  (``revoked_at IS NULL``).
- ``effective_permissions`` — JOIN through the profile's
  ``capabilities`` JSONB and return the capability-name set.
"""

from __future__ import annotations

from datetime import datetime
from uuid import UUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    Any,
    AsyncSessionFactoryPort,
    Sequence,
    sa,
    select,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.assignment_table import (
    PROFILES_TABLE,  # re-exported from profile_table via assignment_table
    USER_APP_ASSIGNMENTS_TABLE,
)
from app.src.modules.lanzadera.domain.assignment import Assignment


def _row_to_assignment(row: sa.Row[Any]) -> Assignment:
    """Map a ``user_app_assignments`` row to the domain ``Assignment`` dataclass."""
    return Assignment(
        id=row.id,
        user_id=row.user_id,
        app_id=row.app_id,
        profile_id=row.profile_id,
        granted_by=row.granted_by,
        granted_at=row.granted_at,
        revoked_at=row.revoked_at,
    )


class AssignmentRepositoryPg:
    """Postgres adapter for the ``AssignmentRepositoryPort`` Protocol (DA-12, H11)."""

    def __init__(self, session_factory: AsyncSessionFactoryPort) -> None:
        self._factory = session_factory

    async def create(self, user_id: UUID, app_id: int, profile_id: UUID) -> Assignment:
        """Persist a new assignment.

        ``id`` and ``granted_at`` are server-defaulted (`gen_random_uuid()`
        / ``NOW()``). The adapter returns the row after a `RETURNING`
        so the domain layer sees the authoritative IDs (DA-1: the
        adapter never invents IDs client-side).

        ``async with self._factory.transaction()`` wraps the INSERT
        AND the 2nd SELECT inside a single transaction: the helper
        commits on exit and rolls back on exception. The 2nd SELECT
        lives inside the transaction (the read-only semantics is
        preserved by virtue of READ COMMITTED isolation -- reads do
        not block writes).
        """
        async with self._factory.transaction() as session:
            stmt = (
                sa.insert(USER_APP_ASSIGNMENTS_TABLE)
                .values(
                    user_id=user_id,
                    app_id=app_id,
                    profile_id=profile_id,
                    granted_by=None,
                )
                .returning(
                    USER_APP_ASSIGNMENTS_TABLE.c.id,
                    USER_APP_ASSIGNMENTS_TABLE.c.granted_at,
                )
            )
            await session.execute(stmt)
            # Re-SELECT with the now-known fields so the returned
            # dataclass carries every attribute. The 2nd read is
            # cheaper than chasing every RETURNING column.
            stmt2 = select(USER_APP_ASSIGNMENTS_TABLE).where(
                sa.and_(
                    USER_APP_ASSIGNMENTS_TABLE.c.user_id == user_id,
                    USER_APP_ASSIGNMENTS_TABLE.c.app_id == app_id,
                    USER_APP_ASSIGNMENTS_TABLE.c.profile_id == profile_id,
                )
            )
            row = (await session.execute(stmt2)).first()
        if row is None:
            raise RuntimeError(
                f"create() returned but row not found for user={user_id} "
                f"app={app_id} profile={profile_id}"
            )
        return _row_to_assignment(row)

    async def list_for_user(self, user_id: UUID) -> Sequence[Assignment]:
        """Return the live assignments for ``user_id`` (``revoked_at IS NULL``)."""
        async with self._factory.read_only_session() as session:
            stmt = (
                select(USER_APP_ASSIGNMENTS_TABLE)
                .where(
                    sa.and_(
                        USER_APP_ASSIGNMENTS_TABLE.c.user_id == user_id,
                        USER_APP_ASSIGNMENTS_TABLE.c.revoked_at.is_(None),
                    )
                )
                .order_by(USER_APP_ASSIGNMENTS_TABLE.c.granted_at.desc())
            )
            rows = (await session.execute(stmt)).all()
        return [_row_to_assignment(r) for r in rows]

    async def list_for_app(self, app_id: int) -> Sequence[Assignment]:
        """Return the live assignments for ``app_id`` (``revoked_at IS NULL``)."""
        async with self._factory.read_only_session() as session:
            stmt = (
                select(USER_APP_ASSIGNMENTS_TABLE)
                .where(
                    sa.and_(
                        USER_APP_ASSIGNMENTS_TABLE.c.app_id == app_id,
                        USER_APP_ASSIGNMENTS_TABLE.c.revoked_at.is_(None),
                    )
                )
                .order_by(USER_APP_ASSIGNMENTS_TABLE.c.user_id)
            )
            rows = (await session.execute(stmt)).all()
        return [_row_to_assignment(r) for r in rows]

    async def effective_permissions(self, user_id: UUID, app_id: int) -> Sequence[str]:
        """Return the capability names from the user's live profile for ``app_id``.

        The query joins ``user_app_assignments`` to ``profiles`` on
        ``profile_id`` and only returns ``active`` rows where the
        assignment's ``revoked_at`` is NULL. The result is the
        flattened ``capabilities`` JSONB keys — the application layer
        treats these as capability names (not values) because the
        platform's permission check is set-membership based
        (DA-12 + H11).
        """
        async with self._factory.read_only_session() as session:
            stmt = (
                select(PROFILES_TABLE.c.capabilities)
                .join(
                    USER_APP_ASSIGNMENTS_TABLE,
                    USER_APP_ASSIGNMENTS_TABLE.c.profile_id == PROFILES_TABLE.c.id,
                )
                .where(
                    sa.and_(
                        USER_APP_ASSIGNMENTS_TABLE.c.user_id == user_id,
                        USER_APP_ASSIGNMENTS_TABLE.c.app_id == app_id,
                        USER_APP_ASSIGNMENTS_TABLE.c.revoked_at.is_(None),
                        PROFILES_TABLE.c.active.is_(True),
                    )
                )
            )
            rows = (await session.execute(stmt)).all()
        capabilities: set[str] = set()
        for (caps,) in rows:
            if caps:
                capabilities.update(caps.keys())
        return sorted(capabilities)

        async def revoke(self, user_id: UUID, app_id: int, *, now: datetime) -> Assignment | None:
            """Soft-delete the live assignment for ``(user_id, app_id)``.

            Sets ``revoked_at = now`` on the row whose ``user_id`` and
            ``app_id`` match and whose ``revoked_at`` is still ``NULL``.
            Returns the updated row, or ``None`` if no live assignment exists.
            """
            async with self._factory.transaction() as session:
                stmt = (
                    sa.update(USER_APP_ASSIGNMENTS_TABLE)
                    .where(
                        sa.and_(
                            USER_APP_ASSIGNMENTS_TABLE.c.user_id == user_id,
                            USER_APP_ASSIGNMENTS_TABLE.c.app_id == app_id,
                            USER_APP_ASSIGNMENTS_TABLE.c.revoked_at.is_(None),
                        )
                    )
                    .values(revoked_at=now)
                    .returning(USER_APP_ASSIGNMENTS_TABLE)
                )
                row = (await session.execute(stmt)).first()
            if row is None:
                return None
            return _row_to_assignment(row)


__all__ = [
    "AssignmentRepositoryPg",
    "PROFILES_TABLE",
    "USER_APP_ASSIGNMENTS_TABLE",
]
