# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#539)
"""Admin use case: login (D38, D39, D89, D-W62-2, DA-11).

Verifies a credential pair, enforces the lockout policy (D38, D39),
creates a fresh ``Session`` row, and emits the audit event for the
operation. Returns the new ``Session``; the caller (delivery, PR-6)
wraps it in a JWT before returning the response.

D38: a user is locked after ``LockoutPolicy.threshold`` consecutive
failed attempts. The lock lasts ``LockoutPolicy.duration_seconds``
unless a global admin intervenes (D39).
D89: the only persisted credential is the password hash (no plaintext
on disk); the reset flow (D90) is the only way to obtain an initial
hash for a PASSWORD_RESET_REQUIRED row.
D-W62-2: sessions persist with the default 24 h TTL; revocation
shortens ``expires_at`` (handled in PR-3).
DA-11: the audit row is appended for both success and failure paths
so the security log captures every attempt.
"""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.errors import (
    AccountLockedError,
    AccountNotActiveError,
    InvalidCredentialsError,
)
from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    AuditLogEntry,
    PasswordHasher,
)
from app.src.modules.lanzadera.domain.ports.session_repository import (
    SessionRepository,
)
from app.src.modules.lanzadera.domain.session import LockoutPolicy, Session
from app.src.modules.lanzadera.domain.user import User, UserStatus

if TYPE_CHECKING:
    from app.src.modules.lanzadera.domain.ports import UserRepository

#: D-W62-2 — session TTL is 24 hours by default. The lockout policy
#: threshold / duration are read from the ``LockoutPolicy`` value object
#: (D38, D39); the TTL itself lives here as a constant so the auth
#: routes can override it without duplicating the value.
DEFAULT_SESSION_TTL_SECONDS = 86_400


def _require_utc(now: datetime) -> None:
    """Reject naive datetimes at the boundary (DA-1 consistency)."""
    from datetime import UTC

    if now.tzinfo is None or now.tzinfo != UTC:
        raise ValueError(f"login requires UTC datetimes; got tzinfo={now.tzinfo!r}")


def _last_failure_within_lockout(user: User, now: datetime, lockout: LockoutPolicy) -> bool:
    """Return True if the user's last login attempt was within the lockout window.

    ``User.failed_attempts`` counts every failure; the time-since-last
    attempt is computed from ``last_login_at`` (recorded on every
    attempt, success or failure). When ``failed_attempts < threshold``
    the user is not locked, regardless of timing.
    """
    if user.failed_attempts < lockout.threshold:
        return False
    if user.last_login_at is None:
        return False
    elapsed = (now - user.last_login_at).total_seconds()
    return elapsed < lockout.duration_seconds


async def login(
    email: str,
    password: str,
    *,
    now: datetime,
    users: "UserRepository",
    sessions: SessionRepository,
    password_hasher: PasswordHasher,
    audit: AuditLog,
    lockout: LockoutPolicy,
    session_ttl_seconds: int = DEFAULT_SESSION_TTL_SECONDS,
    actor_id: UUID | None = None,
) -> Session:
    """Verify the credential pair and return a fresh ``Session``.

    The caller (delivery, PR-6) is responsible for translating the
    typed exceptions to HTTP status codes. The mapping is documented
    on the auth routes in PR-6.

    Raises:
        ``InvalidCredentialsError`` — unknown email or wrong password.
            The two cases are not distinguished so the response time
            cannot be used to enumerate users (D-W62-2, D38).
        ``AccountLockedError`` — ``failed_attempts >= threshold`` AND the
            last attempt was within ``duration_seconds`` (D38, D39).
        ``AccountNotActiveError`` — the row exists and the password
            matches, but ``status`` is not ``ACTIVE``.

    Side effects (in order, DA-11):
        1. ``UserRepository.get_by_email(email)`` — read.
        2. ``PasswordHasher.verify(password, hash)`` — verify.
        3. ``UserRepository.update_failed_attempts(...)`` (failure) or
           ``reset_failed_attempts(...)`` (success) — counter mutation.
        4. ``UserRepository.record_login_attempt(...)`` — last-login stamp.
        5. ``SessionRepository.create(session)`` — session insert.
        6. ``AuditLog.append(...)`` — security log row.
    """
    _require_utc(now)
    if not email or not email.strip():
        raise ValueError("login.email must be a non-empty string")
    if not password:
        raise ValueError("login.password must be a non-empty string")

    normalised_email = email.strip().lower()
    user = await users.get_by_email(normalised_email)

    # Constant-time guard: when the email is unknown we still burn a
    # hash cycle and still append an audit row. The ``target_id``
    # falls back to the email so the security log can correlate
    # enumeration attempts with their payload.
    if user is None:
        await password_hasher.verify(password, "fake:nonexistent")
        await audit.append(
            AuditLogEntry(
                event_type="auth.login.failure",
                actor_id=None,
                target_id=normalised_email,
                result="failure",
                created_at=now,
                module="lanzadera",
                correlation_id=uuid4(),
                payload={"reason": "unknown_email"},
            )
        )
        raise InvalidCredentialsError("invalid credentials")

    # Lockout gate (D38, D39). We check the counter + the timing window
    # together; either the counter is below the threshold or the
    # lockout window has elapsed, the user may attempt again.
    if _last_failure_within_lockout(user, now, lockout):
        await audit.append(
            AuditLogEntry(
                event_type="auth.login.failure",
                actor_id=None,
                target_id=str(user.id),
                result="failure",
                created_at=now,
                module="lanzadera",
                correlation_id=uuid4(),
                payload={"reason": "locked", "failed_attempts": user.failed_attempts},
            )
        )
        raise AccountLockedError(f"account is locked after {user.failed_attempts} failed attempts")

    # Password verification.
    if user.password_hash is None or not await password_hasher.verify(password, user.password_hash):
        new_failed = user.failed_attempts + 1
        await users.update_failed_attempts(user.id, new_failed)
        await users.record_login_attempt(user.id, at=now)
        await audit.append(
            AuditLogEntry(
                event_type="auth.login.failure",
                actor_id=None,
                target_id=str(user.id),
                result="failure",
                created_at=now,
                module="lanzadera",
                correlation_id=uuid4(),
                payload={
                    "reason": "bad_password",
                    "failed_attempts": new_failed,
                },
            )
        )
        raise InvalidCredentialsError("invalid credentials")

    # Status gate (D89). Only ACTIVE users may log in. ``reset_required``
    # and ``disabled`` are out-of-band paths.
    if user.status is not UserStatus.ACTIVE:
        await audit.append(
            AuditLogEntry(
                event_type="auth.login.failure",
                actor_id=None,
                target_id=str(user.id),
                result="failure",
                created_at=now,
                module="lanzadera",
                correlation_id=uuid4(),
                payload={"reason": "not_active", "status": user.status.value},
            )
        )
        raise AccountNotActiveError(f"account is not active; status={user.status.value!r}")

    # Success path (DA-11): reset counter, stamp last-login, create the
    # session row, append the audit event. The order matters — the
    # session row carries the user_id the audit row references.
    await users.reset_failed_attempts(user.id)
    await users.record_login_attempt(user.id, at=now)
    session = Session(
        id=uuid4(),
        user_id=user.id,
        created_at=now,
        expires_at=now + timedelta(seconds=session_ttl_seconds),
    )
    await sessions.create(session)
    await audit.append(
        AuditLogEntry(
            event_type="auth.login.success",
            actor_id=actor_id or user.id,
            target_id=str(session.id),
            result="success",
            created_at=now,
            module="lanzadera",
            correlation_id=uuid4(),
            payload={"user_id": str(user.id)},
        )
    )
    return session


__all__ = ["login", "DEFAULT_SESSION_TTL_SECONDS"]
