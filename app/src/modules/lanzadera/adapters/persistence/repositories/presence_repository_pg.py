# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — Postgres adapter for the presence port.
"""Async Postgres adapter for the ``PresenceRepository`` Protocol.

Implements the contract declared in
``app.src.modules.lanzadera.domain.ports.presence_repository.PresenceRepository``
(W60, #522). The four operations are the slice the SSE real-time
presence feature needs:

- ``track`` — INSERT ... ON CONFLICT (user_id) DO UPDATE SET last_seen = now().
  Upsert so reconnects do not have to clear the row first; ``email`` is
  set or refreshed on every track call so an email change mid-session
  propagates.
- ``heartbeat`` — UPDATE last_seen = now() WHERE user_id. Hot path
  (every ``POST /presence/heartbeat`` from the browser); operates on a
  single row and is a no-op if the user was never ``track``-ed.
- ``disconnect`` — DELETE WHERE user_id. Idempotent (no error if the
  row is already gone). Reserved for future cleanup paths (cron
  pruner that drops rows older than the heartbeat TTL); the HTTP
  layer does not call it yet.
- ``list_connected`` — SELECT ORDER BY last_seen DESC LIMIT. Read-only
  path used by both ``GET /presence`` (one-shot snapshot) and
  ``GET /presence/stream`` (5-second SSE poll).

The adapter is constructed once at boot and shared between the
HTTP delivery routes and the test fixtures (mirrors the DA-1 / D89
contract the rest of the Postgres adapters pin). It does NOT own
long-lived connections: every method opens a fresh ``AsyncSession``
through the injected ``AsyncSessionFactoryPort`` and closes it
before returning.
"""

from __future__ import annotations

from uuid import UUID

from sqlalchemy.dialects.postgresql import insert as pg_insert

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    Any,
    AsyncSessionFactoryPort,
    Sequence,
    sa,
    select,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.presence_table import (
    PRESENCE_TABLE,
)
from app.src.modules.lanzadera.domain.presence import ConnectedUser


def _row_to_connected_user(row: sa.Row[Any]) -> ConnectedUser:
    """Map a ``presence_sessions`` row to the domain ``ConnectedUser`` dataclass.

    Forward-only mapping (DB → domain). The reverse direction lives in
    ``track``, which writes the columns explicitly so the schema stays
    the source of truth.
    """
    return ConnectedUser(
        user_id=row.user_id,
        email=row.email,
        connected_at=row.connected_at,
        last_seen=row.last_seen,
    )


class PresenceRepositoryPg:
    """Postgres adapter for the ``PresenceRepository`` Protocol (W60, #522)."""

    def __init__(self, session_factory: AsyncSessionFactoryPort) -> None:
        self._factory = session_factory

    async def track(self, user_id: UUID, email: str) -> None:
        """Upsert a presence row.

        INSERT ... ON CONFLICT (user_id) DO UPDATE SET last_seen = now().
        The ``email`` column is refreshed on every track call so the SSE
        snapshot reflects the latest known identity; ``connected_at`` is
        preserved by the ``ON CONFLICT`` clause (only ``INSERT`` writes
        ``connected_at``).
        """
        async with self._factory.transaction() as session:
            stmt = pg_insert(PRESENCE_TABLE).values(
                user_id=user_id,
                email=email,
                connected_at=sa.func.now(),
                last_seen=sa.func.now(),
            )
            stmt = stmt.on_conflict_do_update(
                index_elements=[PRESENCE_TABLE.c.user_id],
                set_={"email": stmt.excluded.email, "last_seen": sa.func.now()},
            )
            await session.execute(stmt)

    async def heartbeat(self, user_id: UUID) -> None:
        """Refresh ``last_seen`` for ``user_id`` (no-op if no row exists).

        A non-existent row stays non-existent — the HTTP layer treats a
        heartbeat without a prior track as a "user reconnected mid-session"
        and the next ``track`` call will mint a fresh row.
        """
        async with self._factory.transaction() as session:
            stmt = (
                sa.update(PRESENCE_TABLE)
                .where(PRESENCE_TABLE.c.user_id == user_id)
                .values(last_seen=sa.func.now())
            )
            await session.execute(stmt)

    async def disconnect(self, user_id: UUID) -> None:
        """Delete the presence row for ``user_id`` (idempotent).

        No exception is raised if the row is already gone — the cleanup
        cron that will eventually call this (out of scope for W60) is
        one of many concurrent sources of deletion.
        """
        async with self._factory.transaction() as session:
            stmt = sa.delete(PRESENCE_TABLE).where(PRESENCE_TABLE.c.user_id == user_id)
            await session.execute(stmt)

    async def list_connected(self, limit: int = 100) -> Sequence[ConnectedUser]:
        """Return the most recently active users, ordered by ``last_seen`` DESC.

        ``limit`` is honoured verbatim — the SSE emitter caps the result
        so a busy admin console stays within a single network round-trip.
        Ordering by ``last_seen`` DESC keeps the freshest heartbeat at the
        top so the admin viewer renders the most relevant user first.
        """
        async with self._factory.read_only_session() as session:
            stmt = select(PRESENCE_TABLE).order_by(PRESENCE_TABLE.c.last_seen.desc()).limit(limit)
            rows = (await session.execute(stmt)).all()
        return [_row_to_connected_user(r) for r in rows]


__all__ = ["PresenceRepositoryPg"]
