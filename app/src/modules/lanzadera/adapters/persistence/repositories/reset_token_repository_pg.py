# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp W05 (#45-subset)
# DA-4, D90 — Postgres adapter for the reset-token port.
"""Async Postgres adapter for the ``ResetTokenRepositoryPort`` Protocol.

Implements the contract declared in
``app.src.modules.lanzadera.domain.ports.reset_token_repository.ResetTokenRepositoryPort``
(DA-4, D90). The five-method API:

- ``insert(user_id, token_hash, expires_at)`` — INSERT a fresh row.
  Uniqueness on ``token_hash`` is enforced by the schema's
  ``uq_reset_tokens_token_hash`` index.
- ``find_unused(token_hash)`` — return the row iff
  ``consumed_at IS NULL AND superseded_at IS NULL AND expires_at >
  now()``. Returns ``None`` for unknown / expired / consumed /
  superseded rows.
- ``mark_consumed(token_hash, at)`` — STAMP ``consumed_at``.
  Idempotent on a second call (a no-op when already consumed).
- ``mark_superseded(user_id, at)`` — STAMP ``superseded_at`` on every
  live row for ``user_id``. The D90 supersession rule.
- ``purge_expired(now)`` — DELETE rows with ``expires_at < now``;
  return the deleted count.

The atomicity story (DA-11) is the application's responsibility:
``mark_consumed`` MUST run in the same SQLAlchemy session as the
``UserRepository.update_password_and_activate`` mutation, so a
failed ``AuditLogPort.append`` rolls back the consume. The adapter
opens its own ``AsyncSession`` per call here; transactional
consumers wire the adapter via session sharing instead.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any
from uuid import UUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    AsyncSessionFactoryPort,
    Sequence,
    sa,
    select,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.reset_token_table import (
    RESET_TOKENS_TABLE,
)  # noqa: F401  # re-exported for back-compat
from app.src.modules.lanzadera.domain.reset_token import ResetToken


def _row_to_token(row: sa.Row[Any]) -> ResetToken:
    """Map a ``reset_tokens`` row to the domain ``ResetToken`` dataclass."""
    return ResetToken(
        id=row.id,
        user_id=row.user_id,
        token_hash=row.token_hash,
        expires_at=row.expires_at,
        consumed_at=row.consumed_at,
        superseded_at=row.superseded_at,
        created_at=row.created_at,
    )


class ResetTokenRepositoryPg:
    """Postgres adapter for the ``ResetTokenRepositoryPort`` Protocol (DA-4, D90)."""

    def __init__(self, session_factory: AsyncSessionFactoryPort) -> None:
        self._factory = session_factory

    async def insert(
        self,
        user_id: UUID,
        token_hash: str,
        expires_at: datetime,
    ) -> ResetToken:
        """Persist a fresh token. Uniqueness on ``token_hash`` is enforced."""
        async with self._factory.transaction() as session:
            stmt = (
                sa.insert(RESET_TOKENS_TABLE)
                .values(
                    user_id=user_id,
                    token_hash=token_hash,
                    expires_at=expires_at,
                )
                .returning(
                    RESET_TOKENS_TABLE.c.id,
                    RESET_TOKENS_TABLE.c.created_at,
                )
            )
            await session.execute(stmt)
            # Re-read so the returned dataclass carries server defaults.
            stmt2 = select(RESET_TOKENS_TABLE).where(RESET_TOKENS_TABLE.c.token_hash == token_hash)
            row = (await session.execute(stmt2)).first()
        if row is None:
            raise RuntimeError(f"insert() returned but row not found for hash={token_hash}")
        return _row_to_token(row)

    async def find_unused(self, token_hash: str) -> ResetToken | None:
        """Return the live row, ``None`` for unknown / expired / consumed / superseded."""
        async with self._factory.read_only_session() as session:
            stmt = select(RESET_TOKENS_TABLE).where(
                sa.and_(
                    RESET_TOKENS_TABLE.c.token_hash == token_hash,
                    RESET_TOKENS_TABLE.c.consumed_at.is_(None),
                    RESET_TOKENS_TABLE.c.superseded_at.is_(None),
                    RESET_TOKENS_TABLE.c.expires_at > sa.func.now(),
                )
            )
            row = (await session.execute(stmt)).first()
        return _row_to_token(row) if row is not None else None

    async def mark_consumed(self, token_hash: str, at: datetime) -> None:
        """Stamp ``consumed_at`` once. Idempotent on the second call."""
        async with self._factory.transaction() as session:
            stmt = (
                sa.update(RESET_TOKENS_TABLE)
                .where(RESET_TOKENS_TABLE.c.token_hash == token_hash)
                .values(consumed_at=at)
            )
            await session.execute(stmt)

    async def mark_superseded(self, user_id: UUID, at: datetime) -> None:
        """Stamp ``superseded_at`` on every live row for ``user_id`` (DA-4 / D90)."""
        async with self._factory.transaction() as session:
            stmt = (
                sa.update(RESET_TOKENS_TABLE)
                .where(
                    sa.and_(
                        RESET_TOKENS_TABLE.c.user_id == user_id,
                        RESET_TOKENS_TABLE.c.consumed_at.is_(None),
                        RESET_TOKENS_TABLE.c.superseded_at.is_(None),
                    )
                )
                .values(superseded_at=at)
            )
            await session.execute(stmt)

    async def purge_expired(self, now: datetime) -> int:
        """Delete rows whose ``expires_at < now``; return the deleted count.

        SQLAlchemy 2.0 ``CursorResult.rowcount`` is typed as a plain
        ``int`` on the ``result`` object; mypy sees it but a strict
        stub would mark it ``int | None``. We coerce and default to 0
        on the rare ``None`` case the dialect can produce.
        """
        async with self._factory.transaction() as session:
            stmt = sa.delete(RESET_TOKENS_TABLE).where(RESET_TOKENS_TABLE.c.expires_at < now)
            result = await session.execute(stmt)
        rowcount: int = getattr(result, "rowcount", 0) or 0
        return rowcount


__all__ = [
    "RESET_TOKENS_TABLE",
    "ResetTokenRepositoryPg",
]
