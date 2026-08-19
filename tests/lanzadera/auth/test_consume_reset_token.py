# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 42
# D90 — consume_reset_token is a CRITICAL_HELPER (QC-5).
"""Strict TDD — `consume_reset_token` (D90, DA-4, DA-11)."""

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


def _seed(
    *,
    raw: str = "raw-token-abc",
    expires_offset: timedelta = timedelta(hours=24),
    consumed_at: datetime | None = None,
    superseded_at: datetime | None = None,
    created_at: datetime | None = None,
):
    """Return `(user, users, tokens, audit, hasher)` for one test."""
    from tests.lanzadera.auth._fakes import (  # noqa: PLC0415
        FakeAuditLog,
        FakePasswordHasher,
        FakeResetTokenRepository,
        FakeUserRepository,
    )

    user = _user()
    created = created_at or _now()
    token = ResetToken(
        id=uuid4(),
        user_id=user.id,
        token_hash=hashlib.blake2b(raw.encode("utf-8"), digest_size=32).hexdigest(),
        expires_at=created + expires_offset,
        consumed_at=consumed_at,
        superseded_at=superseded_at,
        created_at=created,
    )
    users = FakeUserRepository()
    users.add(user)
    tokens = FakeResetTokenRepository()
    tokens.add(token)
    return user, users, tokens, FakeAuditLog(), FakePasswordHasher()


def _consume(deps, *, raw: str = "raw-token-abc", password: str = "new-passphrase"):
    _, users, tokens, audit, hasher = deps
    return consume_reset_token(
        raw,
        password,
        now=_now(),
        hasher=hasher,
        users=users,
        reset_tokens=tokens,
        audit=audit,
    )


# -- Happy path --------------------------------------------------------------


def test_returns_none_on_success():
    assert _consume(_seed()) is None


def test_user_password_and_status_updated_with_hashed_value():
    user, *_ = deps = _seed()
    _consume(deps)
    assert user.password_hash == "fake:new-passphrase"
    assert user.status is UserStatus.ACTIVE


def test_token_marked_consumed_and_audit_emitted():
    user, users, tokens, audit, hasher = _seed()
    _consume((user, users, tokens, audit, hasher))
    stored = list(tokens.by_hash.values())[0]
    assert stored.consumed_at == _now()
    assert audit.entries[0].event_type == "auth.reset.consumed"
    assert audit.entries[0].actor_id == user.id


# -- Single-use --------------------------------------------------------------


def test_replay_raises_already_used():
    deps = _seed()
    _consume(deps, password="first")
    with pytest.raises(ResetTokenAlreadyUsedError):
        _consume(deps, password="second")


def test_no_writes_on_replay():
    user, *_ = deps = _seed()
    _consume(deps, password="first")
    snapshot_hash = user.password_hash
    snapshot_audit = len(deps[3].entries)
    try:
        _consume(deps, password="second")
    except ResetTokenAlreadyUsedError:
        pass
    assert user.password_hash == snapshot_hash
    assert len(deps[3].entries) == snapshot_audit


# -- Lifecycle failures ------------------------------------------------------


def test_expired_token_raises_expired_error():
    past = _now() - timedelta(hours=2)
    deps = _seed(created_at=past, expires_offset=timedelta(hours=1))
    with pytest.raises(ExpiredResetTokenError):
        _consume(deps)


def test_superseded_token_raises_invalid_error():
    deps = _seed(superseded_at=_now() - timedelta(seconds=1))
    with pytest.raises(InvalidResetTokenError):
        _consume(deps)


def test_unknown_token_raises_invalid_error():
    deps = _seed()
    with pytest.raises(InvalidResetTokenError):
        _consume(deps, raw="totally-unknown")


# -- Rollback on hash failure -----------------------------------------------


def test_hash_failure_propagates_with_no_writes():
    user, users, tokens, audit, hasher = _seed()
    hasher.fail_next = True
    with pytest.raises(RuntimeError):
        _consume((user, users, tokens, audit, hasher))
    assert user.password_hash is None
    assert user.status is UserStatus.PASSWORD_RESET_REQUIRED
    assert list(tokens.by_hash.values())[0].consumed_at is None
    assert audit.entries == []


def test_audit_failure_propagates():
    user, users, tokens, audit, hasher = _seed()
    audit.next_raises = RuntimeError("audit emission failed")
    with pytest.raises(RuntimeError, match="audit emission failed"):
        _consume((user, users, tokens, audit, hasher))


# -- Atomicity / boundaries --------------------------------------------------


def test_operations_run_in_order_hash_user_audit():
    user, users, tokens, audit, hasher = _seed()
    _consume((user, users, tokens, audit, hasher))
    assert hasher.calls[0][0] == "hash"
    assert len(users.update_calls) == 1
    assert len(audit.entries) == 1


def test_naive_now_is_rejected():
    deps = _seed()
    with pytest.raises(ValueError):
        consume_reset_token(
            "raw-token-abc",
            "new-pass",
            now=datetime(2026, 8, 13, 12, 0, 0),  # noqa: DTZ001
            hasher=deps[4],
            users=deps[1],
            reset_tokens=deps[2],
            audit=deps[3],
        )


def test_empty_password_is_rejected():
    deps = _seed()
    with pytest.raises(ValueError):
        _consume(deps, password="")
