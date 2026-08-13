# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 42
# D90 — consume_reset_token is a CRITICAL_HELPER (QC-5).
"""Strict TDD — `consume_reset_token` (D90, DA-4, DA-11, CRITICAL_HELPER)."""

from __future__ import annotations

import hashlib
from datetime import UTC, datetime, timedelta
from uuid import uuid4

import pytest

from app.src.modules.lanzadera.domain.errors import (
    ExpiredResetTokenError,
    InvalidResetTokenError,
    ResetTokenAlreadyUsedError,
)
from app.src.modules.lanzadera.domain.reset_token import ResetToken
from app.src.modules.lanzadera.domain.services.consume_reset_token import (
    consume_reset_token,
)
from app.src.modules.lanzadera.domain.user import User, UserStatus

from tests.lanzadera.auth._fakes import (
    FakeAuditLog,
    FakePasswordHasher,
    FakeResetTokenRepository,
    FakeUserRepository,
)


def _now() -> datetime:
    return datetime(2026, 8, 13, 12, 0, 0, tzinfo=UTC)


def _user() -> User:
    return User(
        id=uuid4(),
        email="alice@enterprise.test",
        name="Alice",
        dni_encrypted=b"\x00",
        password_hash=None,
        status=UserStatus.PASSWORD_RESET_REQUIRED,
        failed_attempts=0,
        last_login_at=None,
        created_at=_now(),
        updated_at=_now(),
    )


def _seed_token(
    user: User,
    *,
    raw: str = "raw-token-abc",
    expires_at: datetime | None = None,
    consumed_at: datetime | None = None,
    superseded_at: datetime | None = None,
    created_at: datetime | None = None,
) -> tuple[FakeUserRepository, FakeResetTokenRepository, FakeAuditLog, FakePasswordHasher]:
    created = created_at or _now()
    token = ResetToken(
        id=uuid4(),
        user_id=user.id,
        token_hash=hashlib.blake2b(raw.encode("utf-8"), digest_size=32).hexdigest(),
        expires_at=expires_at or (created + timedelta(hours=24)),
        consumed_at=consumed_at,
        superseded_at=superseded_at,
        created_at=created,
    )
    users = FakeUserRepository()
    users.add(user)
    tokens = FakeResetTokenRepository()
    tokens.add(token)
    return users, tokens, FakeAuditLog(), FakePasswordHasher()


# ---------------------------------------------------------------------------
# Happy path
# ---------------------------------------------------------------------------


class TestHappyPath:
    def test_returns_none_on_success(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        result = consume_reset_token(
            "raw-token-abc",
            "new-passphrase",
            now=_now(),
            hasher=hasher,
            users=users,
            reset_tokens=tokens,
            audit=audit,
        )
        assert result is None

    def test_user_password_and_status_updated(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        consume_reset_token(
            "raw-token-abc", "new-passphrase", now=_now(),
            hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
        )
        assert user.password_hash == "fake:new-passphrase"
        assert user.status is UserStatus.ACTIVE

    def test_token_marked_consumed(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        consume_reset_token(
            "raw-token-abc", "new-passphrase", now=_now(),
            hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
        )
        stored = list(tokens.by_hash.values())[0]
        assert stored.consumed_at == _now()

    def test_audit_appended(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        consume_reset_token(
            "raw-token-abc", "new-passphrase", now=_now(),
            hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
        )
        assert len(audit.entries) == 1
        assert audit.entries[0].event_type == "auth.reset.consumed"
        assert audit.entries[0].actor_id == user.id


# ---------------------------------------------------------------------------
# Single-use
# ---------------------------------------------------------------------------


class TestSingleUse:
    def test_replay_raises_already_used(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        consume_reset_token(
            "raw-token-abc", "first-pass", now=_now(),
            hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
        )
        with pytest.raises(ResetTokenAlreadyUsedError):
            consume_reset_token(
                "raw-token-abc", "second-pass", now=_now(),
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )

    def test_no_writes_on_replay(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        consume_reset_token(
            "raw-token-abc", "first-pass", now=_now(),
            hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
        )
        snapshot_hash = user.password_hash
        snapshot_audit = len(audit.entries)
        try:
            consume_reset_token(
                "raw-token-abc", "second-pass", now=_now(),
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )
        except ResetTokenAlreadyUsedError:
            pass
        assert user.password_hash == snapshot_hash
        assert len(audit.entries) == snapshot_audit


# ---------------------------------------------------------------------------
# Expired
# ---------------------------------------------------------------------------


class TestExpired:
    def test_expired_token_raises(self) -> None:
        user = _user()
        past = _now() - timedelta(hours=2)
        users, tokens, audit, hasher = _seed_token(
            user,
            created_at=past,
            expires_at=past + timedelta(hours=1),
        )
        with pytest.raises(ExpiredResetTokenError):
            consume_reset_token(
                "raw-token-abc", "new-pass", now=_now(),
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )

    def test_no_writes_on_expired(self) -> None:
        user = _user()
        past = _now() - timedelta(hours=2)
        users, tokens, audit, hasher = _seed_token(
            user,
            created_at=past,
            expires_at=past + timedelta(hours=1),
        )
        try:
            consume_reset_token(
                "raw-token-abc", "new-pass", now=_now(),
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )
        except ExpiredResetTokenError:
            pass
        assert user.password_hash is None
        assert audit.entries == []


# ---------------------------------------------------------------------------
# Invalid (superseded / unknown)
# ---------------------------------------------------------------------------


class TestInvalid:
    def test_superseded_token_raises_invalid(self) -> None:
        user = _user()
        users = FakeUserRepository()
        users.add(user)
        tokens = FakeResetTokenRepository()
        tokens.add(
            ResetToken(
                id=uuid4(),
                user_id=user.id,
                token_hash=hashlib.blake2b(b"raw-token-abc", digest_size=32).hexdigest(),
                expires_at=_now() + timedelta(hours=24),
                consumed_at=None,
                superseded_at=_now() - timedelta(seconds=1),
                created_at=_now(),
            )
        )
        audit = FakeAuditLog()
        hasher = FakePasswordHasher()
        with pytest.raises(InvalidResetTokenError):
            consume_reset_token(
                "raw-token-abc", "new-pass", now=_now(),
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )

    def test_unknown_token_raises_invalid(self) -> None:
        user = _user()
        users = FakeUserRepository()
        users.add(user)
        tokens = FakeResetTokenRepository()
        audit = FakeAuditLog()
        hasher = FakePasswordHasher()
        with pytest.raises(InvalidResetTokenError):
            consume_reset_token(
                "totally-unknown", "new-pass", now=_now(),
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )


# ---------------------------------------------------------------------------
# Rollback on hash failure
# ---------------------------------------------------------------------------


class TestRollbackOnHashFailure:
    def test_hash_failure_propagates(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        hasher.fail_next = True
        with pytest.raises(RuntimeError, match="simulated hash failure"):
            consume_reset_token(
                "raw-token-abc", "new-pass", now=_now(),
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )

    def test_no_writes_on_hash_failure(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        hasher.fail_next = True
        try:
            consume_reset_token(
                "raw-token-abc", "new-pass", now=_now(),
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )
        except RuntimeError:
            pass
        assert user.password_hash is None
        assert user.status is UserStatus.PASSWORD_RESET_REQUIRED
        assert audit.entries == []
        stored = list(tokens.by_hash.values())[0]
        assert stored.consumed_at is None

    def test_audit_failure_propagates(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        audit.next_raises = RuntimeError("audit emission failed")
        with pytest.raises(RuntimeError, match="audit emission failed"):
            consume_reset_token(
                "raw-token-abc", "new-pass", now=_now(),
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )


# ---------------------------------------------------------------------------
# Atomicity / boundaries
# ---------------------------------------------------------------------------


class TestAtomicity:
    def test_operations_in_order(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        consume_reset_token(
            "raw-token-abc", "new-pass", now=_now(),
            hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
        )
        assert hasher.calls[0][0] == "hash"
        assert len(users.update_calls) == 1
        assert len(audit.entries) == 1

    def test_naive_now_is_rejected(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        with pytest.raises(ValueError):
            consume_reset_token(
                "raw-token-abc", "new-pass",
                now=datetime(2026, 8, 13, 12, 0, 0),  # noqa: DTZ001
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )

    def test_empty_password_is_rejected(self) -> None:
        user = _user()
        users, tokens, audit, hasher = _seed_token(user)
        with pytest.raises(ValueError):
            consume_reset_token(
                "raw-token-abc", "", now=_now(),
                hasher=hasher, users=users, reset_tokens=tokens, audit=audit,
            )
