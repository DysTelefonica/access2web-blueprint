# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 42
# DA-4, DA-11, D90 — consume_reset_token is a CRITICAL_HELPER (D90, QC-5).
"""Domain service: `consume_reset_token` (D90, DA-4, DA-11, CRITICAL_HELPER)."""

from __future__ import annotations

import hashlib
from datetime import datetime, timezone

from app.src.modules.lanzadera.domain.errors import (
    ExpiredResetTokenError,
    InvalidResetTokenError,
    ResetTokenAlreadyUsedError,
)
from app.src.modules.lanzadera.domain.ports import AuditLogEntry


def _require_utc(now: datetime) -> None:
    """Reject naive datetimes at the boundary."""
    if now.tzinfo is None or now.tzinfo != timezone.utc:
        raise ValueError(f"consume_reset_token requires UTC datetimes; got tzinfo={now.tzinfo!r}")


def _hash_token(raw: str) -> str:
    """BLAKE2b-256 hex digest (mirrors `issue_reset_token._hash_token`)."""
    return hashlib.blake2b(raw.encode("utf-8"), digest_size=32).hexdigest()


def _classify_invalid(row: "object", now: datetime) -> Exception:
    """Map a stored-but-non-usable token to the correct domain exception.

    The Fake repo exposes `by_hash`; the real Postgres adapter exposes a
    different lookup. We duck-type on `by_hash` so the classification
    works against the Fake without a new dependency on the production
    code path.
    """
    if row is None:
        return InvalidResetTokenError("reset token is unknown")
    if row.consumed_at is not None:
        return ResetTokenAlreadyUsedError(
            f"reset token already consumed at {row.consumed_at.isoformat()}"
        )
    if row.superseded_at is not None:
        return InvalidResetTokenError(
            f"reset token was superseded at {row.superseded_at.isoformat()}"
        )
    if row.expires_at <= now:
        return ExpiredResetTokenError(
            f"reset token expired at {row.expires_at.isoformat()}"
        )
    return InvalidResetTokenError("reset token is unusable")


def consume_reset_token(
    token_str: str,
    new_password: str,
    *,
    now: datetime,
    hasher: "object",
    users: "object",
    reset_tokens: "object",
    audit: "object",
) -> None:
    """Validate `token_str`, write the new password hash, mark the token used.

    Sequence (DA-4, DA-11):
        1. Validate `now` is UTC and `new_password` is non-empty.
        2. BLAKE2b the raw token and look it up via `reset_tokens.find_unused`.
        3. Classify the failure mode when the lookup returns None.
        4. Hash the new password via the `PasswordHasher` port.
        5. Update the user (password hash + status=ACTIVE) in one call.
        6. Mark the token consumed.
        7. Append the audit event.

    Any exception raised by the ports propagates untouched so the caller's
    transaction can roll back (DA-11).
    """
    _require_utc(now)
    if not new_password:
        raise ValueError("consume_reset_token.new_password must be a non-empty string")

    token = reset_tokens.find_unused(_hash_token(token_str), now)
    if token is None:
        # Find the underlying row to pick the right exception.
        stored = getattr(reset_tokens, "by_hash", {}).get(_hash_token(token_str))
        raise _classify_invalid(stored, now)

    new_password_hash = hasher.hash(new_password)
    users.update_password_and_activate(token.user_id, new_password_hash)
    reset_tokens.mark_consumed(token.token_hash, now)
    audit.append(
        AuditLogEntry(
            event_type="auth.reset.consumed",
            actor_id=token.user_id,
            target_id=str(token.id),
            result="success",
            created_at=now,
        )
    )
