# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 42
# DA-4, D90 — issue_reset_token is a CRITICAL_HELPER (D90, QC-5).
"""Domain service: `issue_reset_token` (D90, DA-4, CRITICAL_HELPER)."""

from __future__ import annotations

import hashlib
import secrets
from datetime import UTC, datetime, timedelta
from uuid import uuid4

from app.src.modules.lanzadera.domain.errors import (
    NoGlobalAdminError,
    UserNotFoundError,
)
from app.src.modules.lanzadera.domain.ports import (
    GlobalAdminRepository,
    NotificationDelivery,
    ResetTokenRepository,
    UserRepository,
)
from app.src.modules.lanzadera.domain.reset_token import ResetToken


def _require_utc(now: datetime) -> None:
    """Reject naive datetimes at the boundary (whole pipeline is UTC)."""
    if now.tzinfo is None or now.tzinfo != UTC:
        raise ValueError(f"issue_reset_token requires UTC datetimes; got tzinfo={now.tzinfo!r}")


def _hash_token(raw: str) -> str:
    """BLAKE2b-256 hex digest of the raw token (storage key, not password).
    BLAKE2b (RFC 7693) is modern, collision-resistant, and not in the
    legacy-hash walker forbidden-symbol list.
    """
    return hashlib.blake2b(raw.encode("utf-8"), digest_size=32).hexdigest()


def issue_reset_token(
    email: str,
    *,
    now: datetime,
    users: UserRepository,
    reset_tokens: ResetTokenRepository,
    global_admins: GlobalAdminRepository,
    notifications: NotificationDelivery,
    ttl: timedelta = timedelta(hours=24),
) -> ResetToken:
    """Issue a one-time reset token. Raises on guard failures.

    D90 contract:
        * No global admin -> raises `NoGlobalAdminError`.
        * Unknown email -> raises `UserNotFoundError`.
        * Re-issuance marks every prior unconsumed, un-superseded token
          for the same user as superseded.
        * `expires_at` is `now + ttl` (default 24 h).
        * The raw token is delivered via `NotificationDelivery.send`; the
          `ResetToken` value object carries the BLAKE2b hash only.
    """
    _require_utc(now)
    if not global_admins.there_is_any():
        raise NoGlobalAdminError(
            "issue_reset_token requires at least one global admin; "
            "bootstrap via CLI set-password (D91)"
        )

    user = users.get_by_email(email)
    if user is None:
        raise UserNotFoundError(f"no user with email {email!r}")

    reset_tokens.mark_superseded(user.id, now)

    raw_token = secrets.token_urlsafe(32)
    token_hash = _hash_token(raw_token)
    expires_at = now + ttl
    if expires_at <= now:
        raise ValueError(f"issue_reset_token.ttl must be positive; got ttl={ttl!r}")

    token = ResetToken(
        id=uuid4(),
        user_id=user.id,
        token_hash=token_hash,
        expires_at=expires_at,
        consumed_at=None,
        superseded_at=None,
        created_at=now,
    )
    reset_tokens.insert(token)
    notifications.send(
        to=user.email,
        subject="Reset your Lanzadera password",
        body=f"A password reset was requested. Use this token: {raw_token}",
    )
    return token
