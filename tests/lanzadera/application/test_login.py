# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#539)
"""Tests Categoría 2 (unit / use case) para ``application/login.py``.

D38: lockout threshold 5, ventana 1h.
D39: la lockout dura ``LockoutPolicy.duration_seconds`` por defecto.
D89: sólo ``UserStatus.ACTIVE`` cruza el umbral de login.
D-W62-2: TTL de sesión 24h (``DEFAULT_SESSION_TTL_SECONDS = 86_400``).
DA-11: audit row en cada path (success + 4 failure modes).
HR-2 de la skill ``lanzadera-testing-strategy``: ningún ``MagicMock``;
se usan los fakes de ``tests/lanzadera/_fakes.py``.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from typing import TYPE_CHECKING
from uuid import uuid4

import pytest

from app.src.modules.lanzadera.application.login import (
    DEFAULT_SESSION_TTL_SECONDS,
    login,
)
from app.src.modules.lanzadera.domain.errors import (
    AccountLockedError,
    AccountNotActiveError,
    InvalidCredentialsError,
)
from app.src.modules.lanzadera.domain.session import LockoutPolicy
from app.src.modules.lanzadera.domain.user import User, UserStatus
from tests.lanzadera._fakes import (
    FakeAuditLog,
    FakePasswordHasher,
    FakeSessionRepository,
    FakeUserRepository,
)

if TYPE_CHECKING:
    pass


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _now() -> datetime:
    """Anchor a deterministic ``now`` per test for easier reasoning."""
    return datetime(2026, 9, 1, 12, 0, 0, tzinfo=UTC)


def _seed_active_user(
    users: FakeUserRepository,
    *,
    email: str = "alice@enterprise.test",
    password: str = "correct-password",
    failed_attempts: int = 0,
    last_login_at: datetime | None = None,
    status: UserStatus = UserStatus.ACTIVE,
) -> User:
    """Seed an ACTIVE user with a known password hash (``fake:<password>``)."""
    user = User(
        id=uuid4(),
        email=email,
        name="Alice",
        dni_encrypted=b"enc:11111111",
        # The FakePasswordHasher returns ``fake:<plaintext>`` from
        # ``hash`` and verifies equality on that literal — the test
        # passes ``password`` and the seed stores ``fake:<password>``.
        password_hash=f"fake:{password}",
        status=status,
        failed_attempts=failed_attempts,
        last_login_at=last_login_at,
        created_at=_now() - timedelta(days=30),
        updated_at=_now() - timedelta(days=30),
    )
    users.add(user)
    return user


# ---------------------------------------------------------------------------
# Success path
# ---------------------------------------------------------------------------


async def test_login_returns_session_when_credentials_match() -> None:
    users = FakeUserRepository()
    sessions = FakeSessionRepository()
    audit = FakeAuditLog()
    hasher = FakePasswordHasher()
    password = "correct-password"
    user = _seed_active_user(users, password=password)

    session = await login(
        user.email,
        password,
        now=_now(),
        users=users,
        sessions=sessions,
        password_hasher=hasher,
        audit=audit,
        lockout=LockoutPolicy(),
    )

    assert session.user_id == user.id
    # D-W62-2: TTL 24h by default.
    assert session.expires_at - session.created_at == timedelta(seconds=DEFAULT_SESSION_TTL_SECONDS)


async def test_login_creates_exactly_one_session() -> None:
    """DA-11 invariant: one login → one session row."""
    users, sessions, audit, hasher = (
        FakeUserRepository(),
        FakeSessionRepository(),
        FakeAuditLog(),
        FakePasswordHasher(),
    )
    password = "correct-password"
    user = _seed_active_user(users, password=password)

    await login(
        user.email,
        password,
        now=_now(),
        users=users,
        sessions=sessions,
        password_hasher=hasher,
        audit=audit,
        lockout=LockoutPolicy(),
    )

    assert len(sessions.create_calls) == 1
    assert len(sessions.sessions) == 1
    assert sessions.revoke_calls == []


async def test_login_appends_audit_success_row() -> None:
    users, sessions, audit, hasher = (
        FakeUserRepository(),
        FakeSessionRepository(),
        FakeAuditLog(),
        FakePasswordHasher(),
    )
    password = "correct-password"
    user = _seed_active_user(users, password=password)
    now = _now()

    await login(
        user.email,
        password,
        now=now,
        users=users,
        sessions=sessions,
        password_hasher=hasher,
        audit=audit,
        lockout=LockoutPolicy(),
    )

    assert len(audit.entries) == 1
    row = audit.entries[0]
    assert row.event_type == "auth.login.success"
    assert row.result == "success"
    # target_id is the session id (str(UUID)); sessions.create_calls[0]
    # is the raw UUID object — compare via str().
    assert row.target_id == str(sessions.create_calls[0])
    assert row.payload["user_id"] == str(user.id)


# ---------------------------------------------------------------------------
# Failure paths
# ---------------------------------------------------------------------------


async def test_login_raises_invalid_credentials_for_unknown_email() -> None:
    users, sessions, audit, hasher = (
        FakeUserRepository(),
        FakeSessionRepository(),
        FakeAuditLog(),
        FakePasswordHasher(),
    )
    now = _now()

    with pytest.raises(InvalidCredentialsError):
        await login(
            "ghost@enterprise.test",
            "anything",
            now=now,
            users=users,
            sessions=sessions,
            password_hasher=hasher,
            audit=audit,
            lockout=LockoutPolicy(),
        )

    # No session created.
    assert sessions.create_calls == []
    # Audit row recorded with reason="unknown_email"; target_id is the
    # normalised email so the security log can correlate enumeration attempts.
    assert len(audit.entries) == 1
    row = audit.entries[0]
    assert row.event_type == "auth.login.failure"
    assert row.result == "failure"
    assert row.target_id == "ghost@enterprise.test"
    assert row.payload["reason"] == "unknown_email"


async def test_login_raises_invalid_credentials_for_wrong_password() -> None:
    users, sessions, audit, hasher = (
        FakeUserRepository(),
        FakeSessionRepository(),
        FakeAuditLog(),
        FakePasswordHasher(),
    )
    user = _seed_active_user(users, password="correct-password")
    now = _now()

    with pytest.raises(InvalidCredentialsError):
        await login(
            user.email,
            "WRONG-password",
            now=now,
            users=users,
            sessions=sessions,
            password_hasher=hasher,
            audit=audit,
            lockout=LockoutPolicy(),
        )

    # failed_attempts incremented from 0 → 1.
    refreshed = users.by_id[user.id]
    assert refreshed.failed_attempts == 1
    # last_login_at stamped on every attempt (success or failure).
    assert refreshed.last_login_at == now
    # Audit row carries reason="bad_password" + the new counter value.
    assert len(audit.entries) == 1
    row = audit.entries[0]
    assert row.payload["reason"] == "bad_password"
    assert row.payload["failed_attempts"] == 1


async def test_login_raises_account_locked_when_failed_attempts_reach_threshold() -> None:
    users, sessions, audit, hasher = (
        FakeUserRepository(),
        FakeSessionRepository(),
        FakeAuditLog(),
        FakePasswordHasher(),
    )
    now = _now()
    # failed_attempts at threshold + last attempt within the lockout window.
    user = _seed_active_user(
        users,
        password="correct-password",
        failed_attempts=5,
        last_login_at=now - timedelta(minutes=10),  # within 1h window
    )

    with pytest.raises(AccountLockedError):
        await login(
            user.email,
            "correct-password",  # even the right password locks
            now=now,
            users=users,
            sessions=sessions,
            password_hasher=hasher,
            audit=audit,
            lockout=LockoutPolicy(),
        )

    # No session created.
    assert sessions.create_calls == []
    # Audit row recorded with reason="locked".
    assert len(audit.entries) == 1
    assert audit.entries[0].payload["reason"] == "locked"


async def test_login_allows_retry_after_lockout_window_expires() -> None:
    """D39: once ``duration_seconds`` elapses since the last attempt,
    the lockout expires and the user may retry."""
    users, sessions, audit, hasher = (
        FakeUserRepository(),
        FakeSessionRepository(),
        FakeAuditLog(),
        FakePasswordHasher(),
    )
    now = _now()
    password = "correct-password"
    # failed_attempts at threshold + last attempt OUTSIDE the lockout window.
    user = _seed_active_user(
        users,
        password=password,
        failed_attempts=5,
        last_login_at=now - timedelta(hours=2),  # > 1h window
    )

    session = await login(
        user.email,
        password,
        now=now,
        users=users,
        sessions=sessions,
        password_hasher=hasher,
        audit=audit,
        lockout=LockoutPolicy(),
    )

    # Login succeeds — the lockout window expired.
    assert session.user_id == user.id
    # failed_attempts reset on success.
    assert users.by_id[user.id].failed_attempts == 0


async def test_login_raises_account_not_active_for_password_reset_required_user() -> None:
    """D89: a row in PASSWORD_RESET_REQUIRED state cannot log in even with
    the right password; the only path is the reset flow (D90)."""
    users, sessions, audit, hasher = (
        FakeUserRepository(),
        FakeSessionRepository(),
        FakeAuditLog(),
        FakePasswordHasher(),
    )
    now = _now()
    user = _seed_active_user(
        users,
        password="correct-password",
        status=UserStatus.PASSWORD_RESET_REQUIRED,
    )

    with pytest.raises(AccountNotActiveError):
        await login(
            user.email,
            "correct-password",
            now=now,
            users=users,
            sessions=sessions,
            password_hasher=hasher,
            audit=audit,
            lockout=LockoutPolicy(),
        )

    # No session created; audit row carries the actual status.
    assert sessions.create_calls == []
    assert len(audit.entries) == 1
    row = audit.entries[0]
    assert row.payload["reason"] == "not_active"
    assert row.payload["status"] == UserStatus.PASSWORD_RESET_REQUIRED.value


async def test_login_resets_failed_attempts_on_success() -> None:
    users, sessions, audit, hasher = (
        FakeUserRepository(),
        FakeSessionRepository(),
        FakeAuditLog(),
        FakePasswordHasher(),
    )
    now = _now()
    password = "correct-password"
    user = _seed_active_user(
        users, password=password, failed_attempts=3, last_login_at=now - timedelta(minutes=5)
    )

    await login(
        user.email,
        password,
        now=now,
        users=users,
        sessions=sessions,
        password_hasher=hasher,
        audit=audit,
        lockout=LockoutPolicy(),
    )

    # Counter reset on success (DA-11: counters track attempts, not history).
    assert users.by_id[user.id].failed_attempts == 0
    # last_login_at stamped to ``now`` on success.
    assert users.by_id[user.id].last_login_at == now


async def test_login_burns_a_hash_cycle_when_email_unknown() -> None:
    """Constant-time guard: when the email is unknown, the use case still
    invokes ``password_hasher.verify`` with a fake hash so the response
    time cannot be used to enumerate users (D-W62-2, D38)."""
    users, sessions, audit, hasher = (
        FakeUserRepository(),
        FakeSessionRepository(),
        FakeAuditLog(),
        FakePasswordHasher(),
    )
    now = _now()

    with pytest.raises(InvalidCredentialsError):
        await login(
            "ghost@enterprise.test",
            "any-password",
            now=now,
            users=users,
            sessions=sessions,
            password_hasher=hasher,
            audit=audit,
            lockout=LockoutPolicy(),
        )

    # One verify call recorded even though no real user was looked up.
    assert len(hasher.verify_calls) == 1
    # The recorded call is against a known-bad hash; the exact hash is
    # not part of the contract — only that a call was made.
    _plaintext, hashed = hasher.verify_calls[0]
    assert _plaintext == "any-password"
    assert hashed == "fake:nonexistent"
