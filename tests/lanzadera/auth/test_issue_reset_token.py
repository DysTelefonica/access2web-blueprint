# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 42
# D90 — issue_reset_token is a CRITICAL_HELPER (QC-5).
"""Strict TDD — `issue_reset_token` (D90, DA-4, CRITICAL_HELPER)."""

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

from tests.lanzadera.auth._fakes import (
    FakeGlobalAdminRepository,
    FakeNotificationDelivery,
    FakeResetTokenRepository,
    FakeUserRepository,
)


def _now() -> datetime:
    return datetime(2026, 8, 13, 12, 0, 0, tzinfo=UTC)


def _user(email: str = "alice@enterprise.test") -> User:
    return User(
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


def _deps(user: User | None = None, *, has_admin: bool = True):
    users = FakeUserRepository()
    if user is not None:
        users.add(user)
    return {
        "users": users,
        "reset_tokens": FakeResetTokenRepository(),
        "global_admins": FakeGlobalAdminRepository(has_any=has_admin),
        "notifications": FakeNotificationDelivery(),
    }


def _issue(deps: dict, email: str = "alice@enterprise.test", **overrides):
    return issue_reset_token(email, now=_now(), **deps, **overrides)


# Happy path ----------------------------------------------------------------


class TestHappyPath:
    def test_returns_reset_token(self) -> None:
        result = _issue(_deps(user=_user()))
        assert result.user_id is not None
        assert result.token_hash and len(result.token_hash) == 64
        assert result.consumed_at is None and result.superseded_at is None

    def test_persists_with_blake2b_hash(self) -> None:
        deps = _deps(user=_user())
        _issue(deps)
        raw = deps["notifications"].sent[0][2].split("token: ", 1)[1].strip()
        expected = hashlib.blake2b(raw.encode("utf-8"), digest_size=32).hexdigest()
        assert expected in deps["reset_tokens"].by_hash


# TTL -----------------------------------------------------------------------


class TestTTL:
    def test_default_ttl_is_24_hours(self) -> None:
        result = _issue(_deps(user=_user()))
        assert result.expires_at == _now() + timedelta(hours=24)
        assert result.created_at == _now()

    def test_custom_ttl_is_honoured(self) -> None:
        deps = _deps(user=_user())
        result = _issue(deps, ttl=timedelta(minutes=5))
        assert result.expires_at == _now() + timedelta(minutes=5)


# Supersession --------------------------------------------------------------


class TestSupersession:
    def test_marks_prior_unconsumed_tokens_as_superseded(self) -> None:
        deps = _deps(user=_user())
        first = _issue(deps)
        _issue(deps)
        assert deps["reset_tokens"].by_hash[first.token_hash].superseded_at == _now()

    def test_only_targets_same_user(self) -> None:
        alice, bob = _user("alice@enterprise.test"), _user("bob@enterprise.test")
        deps = _deps(user=alice)
        deps["users"].add(bob)
        bob_first = _issue(deps, "bob@enterprise.test")
        _issue(deps)
        assert deps["reset_tokens"].by_hash[bob_first.token_hash].superseded_at is None


# Guards --------------------------------------------------------------------


class TestGuards:
    def test_no_global_admin_raises(self) -> None:
        deps = _deps(user=_user(), has_admin=False)
        with pytest.raises(NoGlobalAdminError):
            _issue(deps)

    def test_unknown_email_raises_user_not_found(self) -> None:
        deps = _deps(user=_user())
        with pytest.raises(UserNotFoundError):
            _issue(deps, "ghost@enterprise.test")

    def test_no_persistence_when_guard_fires(self) -> None:
        deps = _deps(user=_user(), has_admin=False)
        with pytest.raises(NoGlobalAdminError):
            _issue(deps)
        assert deps["reset_tokens"].by_hash == {}
        assert deps["notifications"].sent == []


# UTC clock + notification --------------------------------------------------


class TestBoundary:
    def test_naive_datetime_is_rejected(self) -> None:
        deps = _deps(user=_user())
        with pytest.raises(ValueError):
            issue_reset_token(
                "alice@enterprise.test",
                now=datetime(2026, 8, 13, 12, 0, 0),  # noqa: DTZ001
                **deps,
            )

    def test_notification_sent_exactly_once(self) -> None:
        deps = _deps(user=_user())
        _issue(deps)
        assert len(deps["notifications"].sent) == 1
        to, subject, _ = deps["notifications"].sent[0]
        assert to == "alice@enterprise.test"
        assert "reset" in subject.lower()
