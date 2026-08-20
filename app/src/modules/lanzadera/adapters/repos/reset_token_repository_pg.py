"""PostgreSQL adapter for :class:`ResetTokenRepositoryPort`.

D90 + DA-4: one-time 24h-TTL tokens, atomicity via the SQLAlchemy
session so a failed audit insert rolls back the consume.

**Skeleton** — the bodies raise ``NotImplementedError`` until migration
``0001_core_schema`` (already on main, declares ``reset_tokens``) is
extended with the audit+consume constraints the application layer needs.
The contract is declared today so ``consume_reset_token`` can pin its
tests against ``typing.Protocol`` structural conformance.
"""

from __future__ import annotations

from datetime import datetime
from uuid import UUID

from app.src.modules.lanzadera.domain.ports.reset_token_repository import (
    ResetTokenRepositoryPort,
)
from app.src.modules.lanzadera.domain.reset_token import ResetToken


class ResetTokenRepositoryPg(ResetTokenRepositoryPort):
    """SQLAlchemy Core + asyncpg implementation. Skeleton — see module docstring."""

    async def insert(self, user_id: UUID, token_hash: str, expires_at: datetime) -> ResetToken:
        raise NotImplementedError(
            "ResetTokenRepositoryPg.insert awaits audit column wiring on reset_tokens"
        )

    async def find_unused(self, token_hash: str) -> ResetToken | None:
        raise NotImplementedError(
            "ResetTokenRepositoryPg.find_unused awaits audit column wiring on reset_tokens"
        )

    async def mark_consumed(self, token_hash: str, at: datetime) -> None:
        raise NotImplementedError(
            "ResetTokenRepositoryPg.mark_consumed awaits audit column wiring on reset_tokens"
        )

    async def mark_superseded(self, user_id: UUID, at: datetime) -> None:
        raise NotImplementedError(
            "ResetTokenRepositoryPg.mark_superseded awaits audit column wiring on reset_tokens"
        )

    async def purge_expired(self, now: datetime) -> int:
        raise NotImplementedError(
            "ResetTokenRepositoryPg.purge_expired awaits audit column wiring on reset_tokens"
        )
