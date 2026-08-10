# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, DA-4, D90 — pure domain value object; only stdlib imports allowed.
"""Reset-token record (lanzadera-mvp/auth-reset).

DA-4 + D90: tokens are single-use, 24-hour-TTL, supersedable. Only the
`token_hash` is persisted — the raw value is emitted once via
`NotificationDeliveryPort.send` and never reconstructed. `ResetToken` is
frozen because every mutation must produce a new row; in-place edits
would let a consumed token come back from the dead.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID


@dataclass(frozen=True)
class ResetToken:
    """One-time reset token (frozen value object).

    `consumed_at` and `superseded_at` are mutually exclusive in practice,
    but the dataclass accepts both as None-or-datetime so the boundary can
    set whichever applies. The transition rules live in the application
    layer (PR 4 — `issue_reset_token` / `consume_reset_token`).
    """

    id: UUID
    user_id: UUID
    token_hash: str
    expires_at: datetime
    consumed_at: datetime | None
    superseded_at: datetime | None
    created_at: datetime

    def __post_init__(self) -> None:
        if not self.token_hash:
            raise ValueError("ResetToken.token_hash must be a non-empty string")
        if self.expires_at <= self.created_at:
            raise ValueError(
                f"ResetToken.expires_at must be strictly after created_at; "
                f"got expires_at={self.expires_at!r} created_at={self.created_at!r}"
            )
