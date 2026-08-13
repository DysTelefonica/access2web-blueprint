# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 42
# D90 — issue_reset_token is a CRITICAL_HELPER (QC-5).
"""Strict TDD — `issue_reset_token` (D90, DA-4)."""

from __future__ import annotations

import hashlib
from datetime import UTC, datetime, timedelta
from uuid import uuid4

import pytest

from app.src.modules.lanzadera.domain.errors import (
    NoGlobalAdminError,
    UserNotFoundError,
)
from app.src.modules.lanzadera.domain.services.issue_reset_token import (
    issue_reset_token,
)
from app.src.modules.lanzadera.domain.user import User, UserStatus


def _now() -> datetime:
    return datetime(2026, 8, 13, 12, 0, 0, tzinfo=UTC)


def _setup(email: str = "alice@enterprise.test", *, has_admin: bool = True):
    """Build the four fakes and a user. Caller passes them as kwargs."""
    from tests.lanzadera.auth._fakes import (  # noqa: PLC0415  local in helper
        FakeGlobalAdminRepository,
        FakeNotificationDelivery,
        FakeResetTokenRepository,
        FakeUserRepository,
    )
    user = User(
        id=uuid4(),
        email=email,
        name="Alice",
        dni_encrypted=b"\x00",
        password_hash=None,
        status=UserStatus.PASSWORD_RESET_REQUIRED,
        failed_attempts=0,
        last_login_at=None,
        created_at=_now(),
        updated_at=_now(),
    )
    users = FakeUserRepository()
    users.add(user)
    return {
        "users": users,
        "reset_tokens": FakeResetTokenRepository(),
        "global_admins": FakeGlobalAdminRepository(has_any=has_admin),
        "notifications": FakeNotificationDelivery(),
    }


def _issue(deps, email: str = "alice@enterprise.test", **overrides):
    return issue_reset_token(email, now=_now(), **deps, **overrides)


# -- Happy path --------------------------------------------------------------


def test_returns_reset_token_with_24h_hash():
    result = _issue(_setup())
    assert result.token_hash and len(result.token_hash) == 64
    assert result.user_id is not None
    assert result.consumed_at is None and result.superseded_at is None


def test_persists_with_blake2b_hash_of_raw_token():
    deps = _setup()
    _issue(deps)
    raw = deps["notifications"].sent[0][2].split("token: ", 1)[1].strip()
    expected = hashlib.blake2b(raw.encode("utf-8"), digest_size=32).hexdigest()
    assert expected in deps["reset_tokens"].by_hash


# -- TTL ---------------------------------------------------------------------


def test_default_ttl_is_24_hours():
    result = _issue(_setup())
    assert result.expires_at == _now() + timedelta(hours=24)


def test_custom_ttl_is_honoured():
    deps = _setup()
    result = _issue(deps, ttl=timedelta(minutes=5))
    assert result.expires_at == _now() + timedelta(minutes=5)


# -- Supersession ------------------------------------------------------------


def test_prior_unconsumed_tokens_are_marked_superseded():
    deps = _setup()
    first = _issue(deps)
    _issue(deps)
    assert deps["reset_tokens"].by_hash[first.token_hash].superseded_at == _now()


def test_supersession_does_not_cross_users():
    alice_deps = _setup("alice@enterprise.test")
    _setup("bob@enterprise.test")
    alice_deps["users"].add(
        User(
            id=uuid4(), email="bob@enterprise.test", name="Bob",
            dni_encrypted=b"\x00", password_hash=None,
            status=UserStatus.PASSWORD_RESET_REQUIRED, failed_attempts=0,
            last_login_at=None, created_at=_now(), updated_at=_now(),
        )
    )
    bob_first = _issue(alice_deps, "bob@enterprise.test")
    _issue(alice_deps)
    assert alice_deps["reset_tokens"].by_hash[bob_first.token_hash].superseded_at is None


# -- Guards ------------------------------------------------------------------


def test_no_global_admin_raises_no_global_admin_error():
    with pytest.raises(NoGlobalAdminError):
        _issue(_setup(has_admin=False))


def test_unknown_email_raises_user_not_found():
    with pytest.raises(UserNotFoundError):
        _issue(_setup(), "ghost@enterprise.test")


def test_no_persistence_when_guard_fires():
    deps = _setup(has_admin=False)
    with pytest.raises(NoGlobalAdminError):
        _issue(deps)
    assert deps["reset_tokens"].by_hash == {}
    assert deps["notifications"].sent == []


# -- Boundary --------------------------------------------------------------


def test_naive_datetime_is_rejected():
    deps = _setup()
    with pytest.raises(ValueError):
        issue_reset_token(
            "alice@enterprise.test",
            now=datetime(2026, 8, 13, 12, 0, 0),  # noqa: DTZ001
            **deps,
        )


def test_notification_sent_exactly_once_with_reset_subject():
    deps = _setup()
    _issue(deps)
    assert len(deps["notifications"].sent) == 1
    to, subject, _ = deps["notifications"].sent[0]
    assert to == "alice@enterprise.test"
    assert "reset" in subject.lower()
