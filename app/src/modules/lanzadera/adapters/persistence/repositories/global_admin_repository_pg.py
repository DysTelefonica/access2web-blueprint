# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp W04 (#21)
# DA-1, D21, D42 — Postgres adapter for the global-admin port.
"""Async Postgres adapter for the ``GlobalAdminRepositoryPort`` Protocol.

Implements the contract declared in
``app.src.modules.lanzadera.domain.ports.global_admin_repository.GlobalAdminRepositoryPort``
(DA-1, D21, D42). The presence-or-absence of a row in ``global_admins``
is the sole signal the auth layer trusts — there is no
``is_global_admin`` column on ``users`` and the D42 invariant (at
least one global admin must always exist) is enforced by the adapter
inside ``revoke``, not by an application-layer guard.
"""

from __future__ import annotations

from uuid import UUID

from sqlalchemy.dialects.postgresql import UUID as PGUUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    SCHEMA,
    Any,
    AsyncSessionFactoryPort,
    Sequence,
    sa,
    select,
)
from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin

GLOBAL_ADMINS_TABLE = sa.Table(
    "global_admins",
    sa.MetaData(),
    sa.Column("user_id", PGUUID(as_uuid=True), primary_key=True),
    sa.Column("granted_at", sa.DateTime(timezone=True), nullable=False),
    sa.Column("granted_by", PGUUID(as_uuid=True), nullable=True),
    schema=SCHEMA,
)


def _row_to_admin(row: sa.Row[Any]) -> GlobalAdmin:
    """Map a ``global_admins`` row to the domain ``GlobalAdmin`` dataclass."""
    return GlobalAdmin(user_id=row.user_id)


class GlobalAdminRepositoryPg:
    """Postgres adapter for the ``GlobalAdminRepositoryPort`` Protocol (DA-1, D21, D42)."""

    def __init__(self, session_factory: AsyncSessionFactoryPort) -> None:
        self._factory = session_factory

    async def list_all(self) -> Sequence[GlobalAdmin]:
        """Return every global-admin row (no soft-delete on this table)."""
        async with self._factory.read_only_session() as session:
            stmt = select(GLOBAL_ADMINS_TABLE.c.user_id).order_by(GLOBAL_ADMINS_TABLE.c.granted_at)
            rows = (await session.execute(stmt)).all()
        return [_row_to_admin(r) for r in rows]

    async def is_global_admin(self, user_id: UUID) -> bool:
        """Return True iff a row exists for ``user_id``."""
        async with self._factory.read_only_session() as session:
            stmt = select(sa.literal(1)).where(GLOBAL_ADMINS_TABLE.c.user_id == user_id).limit(1)
            row = (await session.execute(stmt)).first()
        return row is not None

    async def grant(self, user_id: UUID) -> None:
        """Insert a global-admin row for ``user_id``.

        Raises ``ValueError`` if the row already exists (the duplicate
        is detected at the unique-violation round-trip rather than at
        the application layer).
        """
        async with self._factory.transaction() as session:
            stmt = sa.insert(GLOBAL_ADMINS_TABLE).values(user_id=user_id)
            await session.execute(stmt)

    async def revoke(self, user_id: UUID) -> None:
        """Delete the global-admin row for ``user_id``.

        D42: refuse the deletion if it would leave the system without
        a global admin. The check is the adapter's job because the
        invariant is a database invariant.

        The SELECT count + DELETE run inside one
        ``transaction()`` block so a concurrent grant cannot slip
        between the count and the delete (the original
        implementation had the same atomicity guarantee via the
        manual ``try/except/finally``; the helper makes it explicit).
        """
        async with self._factory.transaction() as session:
            count_stmt = select(sa.func.count()).select_from(GLOBAL_ADMINS_TABLE)
            total = (await session.execute(count_stmt)).scalar_one()
            if total <= 1:
                raise ValueError("cannot revoke the last global admin (D42 invariant)")

            stmt = sa.delete(GLOBAL_ADMINS_TABLE).where(GLOBAL_ADMINS_TABLE.c.user_id == user_id)
            await session.execute(stmt)


__all__ = [
    "GLOBAL_ADMINS_TABLE",
    "GlobalAdminRepositoryPg",
]
