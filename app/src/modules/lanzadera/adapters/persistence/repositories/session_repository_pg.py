# DA-1, D-W62-2 — Postgres adapter for the session port.
"""Async Postgres adapter for the ``SessionRepositoryPort`` Protocol (W62, #921).

- ``create(session)`` — INSERT the row verbatim and return it.
- ``get_by_id(session_id)`` — return the row regardless of expiry; the
  auth middleware owns the expiry check (DA-11).
- ``revoke(session_id, at)`` — ``expires_at = LEAST(expires_at, at)``.
  A second revocation never extends an earlier expiry. Returns ``False``
  when no such session exists.

Every method opens its own ``AsyncSession`` through the factory, the same
way the other Postgres adapters do.
"""

from __future__ import annotations

from datetime import datetime
from uuid import UUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    Any,
    AsyncSessionFactoryPort,
    sa,
    select,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.session_table import (
    SESSIONS_TABLE,
)
from app.src.modules.lanzadera.domain.session import Session


def _row_to_session(row: sa.Row[Any]) -> Session:
    """Map a ``sessions`` row to the domain ``Session`` entity."""
    return Session(
        id=row.id,
        user_id=row.user_id,
        created_at=row.created_at,
        expires_at=row.expires_at,
    )


class SessionRepositoryPg:
    """Postgres adapter for the ``SessionRepositoryPort`` Protocol (W62, #921)."""

    def __init__(self, session_factory: AsyncSessionFactoryPort) -> None:
        self._factory = session_factory

    async def create(self, session: Session) -> Session:
        """Persist the session row as given."""
        async with self._factory.transaction() as db:
            await db.execute(
                sa.insert(SESSIONS_TABLE).values(
                    id=session.id,
                    user_id=session.user_id,
                    created_at=session.created_at,
                    expires_at=session.expires_at,
                )
            )
        return session

    async def get_by_id(self, session_id: UUID) -> Session | None:
        """Return the row iff it exists, whatever its expiry."""
        async with self._factory.read_only_session() as db:
            stmt = select(SESSIONS_TABLE).where(SESSIONS_TABLE.c.id == session_id)
            row = (await db.execute(stmt)).first()
        return _row_to_session(row) if row is not None else None

    async def revoke(self, session_id: UUID, at: datetime) -> bool:
        """Shorten ``expires_at`` to ``at`` unless it is already earlier."""
        async with self._factory.transaction() as db:
            stmt = (
                sa.update(SESSIONS_TABLE)
                .where(SESSIONS_TABLE.c.id == session_id)
                .values(expires_at=sa.func.least(SESSIONS_TABLE.c.expires_at, at))
            )
            result = await db.execute(stmt)
            # SQLAlchemy 2.0 types ``execute()`` as ``Result[Any]``; the
            # sync ``CursorResult.rowcount`` lives on the underlying
            # ``CursorResult``. ``getattr`` keeps mypy happy.
            rowcount: int = getattr(result, "rowcount", 0) or 0
        return rowcount > 0
