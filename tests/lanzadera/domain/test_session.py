# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, D38, D39, D40 — `Session` is a mutable entity (lifecycle transitions
# revoke / expire); `LockoutPolicy` is a frozen value object that the auth
# use case reads at boot. Defaults are 5 failures / 3600 s (D38 + D39).
"""Strict TDD — `Session` entity + `LockoutPolicy` value object (Phase 1, task 1.9)."""

from __future__ import annotations

import dataclasses
from datetime import UTC, datetime, timedelta
from uuid import uuid4

import pytest

from app.src.modules.lanzadera.domain.session import LockoutPolicy, Session

# ---------------------------------------------------------------------------
# Helpers / fixtures
# ---------------------------------------------------------------------------


def _now() -> datetime:
    return datetime(2026, 8, 9, 12, 0, 0, tzinfo=UTC)


def _new_session(
    *,
    user_id: object = None,
    created_at: datetime | None = None,
    expires_at: datetime | None = None,
) -> Session:
    now = created_at or _now()
    return Session(
        id=uuid4(),
        user_id=user_id or uuid4(),
        created_at=now,
        expires_at=expires_at or (now + timedelta(hours=12)),
    )


# ---------------------------------------------------------------------------
# Session — construction
# ---------------------------------------------------------------------------


class TestSessionConstruction:
    def test_session_carries_required_attributes(self) -> None:
        session_id = uuid4()
        user_id = uuid4()
        created = _now()
        expires = created + timedelta(hours=12)
        session = Session(
            id=session_id,
            user_id=user_id,
            created_at=created,
            expires_at=expires,
        )
        assert session.id == session_id
        assert session.user_id == user_id
        assert session.created_at == created
        assert session.expires_at == expires

    def test_session_is_mutable(self) -> None:
        """Sessions can be revoked by setting `expires_at` to a past timestamp."""
        session = _new_session()
        session.expires_at = _now() - timedelta(seconds=1)
        assert session.expires_at < _now()

    def test_two_sessions_same_id_compare_equal(self) -> None:
        session_id = uuid4()
        user_id = uuid4()
        created = _now()
        expires = created + timedelta(hours=12)
        a = Session(
            id=session_id, user_id=user_id, created_at=created, expires_at=expires
        )
        b = Session(
            id=session_id, user_id=user_id, created_at=created, expires_at=expires
        )
        assert a == b


# ---------------------------------------------------------------------------
# LockoutPolicy — frozen value object with D38/D39 defaults
# ---------------------------------------------------------------------------


class TestLockoutPolicyDefaults:
    def test_default_threshold_is_5(self) -> None:
        policy = LockoutPolicy()
        assert policy.threshold == 5

    def test_default_duration_is_3600_seconds(self) -> None:
        policy = LockoutPolicy()
        assert policy.duration_seconds == 3600


class TestLockoutPolicyConstruction:
    def test_custom_threshold_is_accepted(self) -> None:
        policy = LockoutPolicy(threshold=10, duration_seconds=1800)
        assert policy.threshold == 10
        assert policy.duration_seconds == 1800

    def test_negative_threshold_is_rejected(self) -> None:
        with pytest.raises(ValueError, match="threshold"):
            LockoutPolicy(threshold=-1)

    def test_zero_threshold_is_rejected(self) -> None:
        """A zero threshold would disable lockout — refused by the invariant."""
        with pytest.raises(ValueError, match="threshold"):
            LockoutPolicy(threshold=0)

    def test_non_positive_duration_is_rejected(self) -> None:
        with pytest.raises(ValueError, match="duration"):
            LockoutPolicy(duration_seconds=0)

    def test_negative_duration_is_rejected(self) -> None:
        with pytest.raises(ValueError, match="duration"):
            LockoutPolicy(duration_seconds=-1)


class TestLockoutPolicyImmutability:
    def test_lockout_policy_is_frozen(self) -> None:
        policy = LockoutPolicy()
        with pytest.raises(dataclasses.FrozenInstanceError):  # type: ignore[misc]
            policy.threshold = 99  # type: ignore[misc]

    def test_equal_policies_compare_equal(self) -> None:
        a = LockoutPolicy(threshold=5, duration_seconds=3600)
        b = LockoutPolicy(threshold=5, duration_seconds=3600)
        assert a == b
