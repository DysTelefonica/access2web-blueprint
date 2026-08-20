"""ResetTokenRepositoryPort.

D90, DA-4. The storage boundary for the one-time password-reset token
flow. The raw token is delivered exactly once via the notification
port and never reconstructed; only the BLAKE2b-256 hash is persisted.
The application layer enforces atomicity (DA-11): ``mark_consumed`` runs
in the same SQLAlchemy session as the user-password mutation, so a
failed ``AuditLogPort.append`` rolls back the consume.

DA-4 sets the three correctness rules the adapter must satisfy:
single-use (a consumed token returns ``None`` on lookup), bounded-TTL
(expired tokens return ``None`` and become eligible for ``purge_expired``),
and supersession (issuing a fresh token invalidates every prior live
token for the same user).
"""

from __future__ import annotations

from datetime import datetime
from typing import Protocol
from uuid import UUID

from app.src.modules.lanzadera.domain.reset_token import ResetToken


class ResetTokenRepositoryPort(Protocol):
    """Persistence boundary for one-time reset tokens."""

    async def insert(self, user_id: UUID, token_hash: str, expires_at: datetime) -> ResetToken:
        """Persist a fresh token. Uniqueness on ``token_hash`` is enforced."""
        ...

    async def find_unused(self, token_hash: str) -> ResetToken | None:
        """Return the token iff ``consumed_at IS NULL AND superseded_at IS NULL``.

        Expired tokens (``expires_at < now``) are returned as ``None``
        and should be purged by the caller.
        """
        ...

    async def mark_consumed(self, token_hash: str, at: datetime) -> None:
        """Stamp ``consumed_at``. Idempotent on the second call (DA-4)."""
        ...

    async def mark_superseded(self, user_id: UUID, at: datetime) -> None:
        """Mark every live token for ``user_id`` as superseded (DA-4)."""
        ...

    async def purge_expired(self, now: datetime) -> int:
        """Delete rows with ``expires_at < now`` and no audit-bearing state.

        Returns the number of rows deleted (for the operator harness).
        """
        ...
