# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, DA-4, D90 — `ResetToken` is a value object (frozen). The raw token is
# never stored — only its hash. `expires_at` is +24h from creation by
# convention; `consumed_at` and `superseded_at` are mutually exclusive.
"""Strict TDD — `ResetToken` value object (Phase 1, task 1.5)."""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from uuid import uuid4

import pytest

from app.src.modules.lanzadera.domain.reset_token import ResetToken

# ---------------------------------------------------------------------------
# Helpers / fixtures
# ---------------------------------------------------------------------------


def _now() -> datetime:
    return datetime(2026, 8, 9, 12, 0, 0, tzinfo=UTC)


def _new_token(
    *,
    user_id: object = None,
    token_hash: str = "h" * 64,
    expires_at: datetime | None = None,
    consumed_at: datetime | None = None,
    superseded_at: datetime | None = None,
) -> ResetToken:
    return ResetToken(
        id=uuid4(),
        user_id=user_id or uuid4(),
        token_hash=token_hash,
        expires_at=expires_at or (_now() + timedelta(hours=24)),
        consumed_at=consumed_at,
        superseded_at=superseded_at,
        created_at=_now(),
    )


# ---------------------------------------------------------------------------
# Construction
# ---------------------------------------------------------------------------


class TestResetTokenConstruction:
    def test_token_carries_required_attributes(self) -> None:
        token_id = uuid4()
        user_id = uuid4()
        expires_at = _now() + timedelta(hours=24)
        token = ResetToken(
            id=token_id,
            user_id=user_id,
            token_hash="a" * 64,
            expires_at=expires_at,
            consumed_at=None,
            superseded_at=None,
            created_at=_now(),
        )
        assert token.id == token_id
        assert token.user_id == user_id
        assert token.token_hash == "a" * 64
        assert token.expires_at == expires_at
        assert token.consumed_at is None
        assert token.superseded_at is None

    def test_token_hash_must_be_non_empty(self) -> None:
        with pytest.raises(ValueError, match="token_hash"):
            _new_token(token_hash="")

    def test_expires_at_must_be_after_created_at(self) -> None:
        """D90: tokens MUST expire — a non-positive window makes the row invalid."""
        created = _now()
        with pytest.raises(ValueError, match="expires_at"):
            ResetToken(
                id=uuid4(),
                user_id=uuid4(),
                token_hash="x" * 64,
                expires_at=created,  # not strictly after created_at
                consumed_at=None,
                superseded_at=None,
                created_at=created,
            )


# ---------------------------------------------------------------------------
# Immutability — value object
# ---------------------------------------------------------------------------


class TestResetTokenImmutability:
    def test_token_is_frozen(self) -> None:
        """DA-4: tokens are single-use value objects; assignment is forbidden."""
        token = _new_token()
        with pytest.raises(dataclasses.FrozenInstanceError):  # type: ignore[misc]
            token.token_hash = "y" * 64  # type: ignore[misc]

    def test_equal_tokens_compare_equal(self) -> None:
        token_id = uuid4()
        user_id = uuid4()
        expires_at = _now() + timedelta(hours=24)
        created = _now()
        a = ResetToken(
            id=token_id,
            user_id=user_id,
            token_hash="z" * 64,
            expires_at=expires_at,
            consumed_at=None,
            superseded_at=None,
            created_at=created,
        )
        b = ResetToken(
            id=token_id,
            user_id=user_id,
            token_hash="z" * 64,
            expires_at=expires_at,
            consumed_at=None,
            superseded_at=None,
            created_at=created,
        )
        assert a == b


# ---------------------------------------------------------------------------
# Lifecycle flags
# ---------------------------------------------------------------------------


class TestResetTokenLifecycle:
    def test_consumed_token_records_timestamp(self) -> None:
        consumed_at = _now() + timedelta(hours=1)
        token = _new_token(consumed_at=consumed_at)
        assert token.consumed_at == consumed_at

    def test_superseded_token_records_timestamp(self) -> None:
        """DA-4: re-issuing a token supersedes the previous one."""
        superseded_at = _now() + timedelta(hours=1)
        token = _new_token(superseded_at=superseded_at)
        assert token.superseded_at == superseded_at

    def test_active_token_has_no_consumed_or_superseded(self) -> None:
        token = _new_token()
        assert token.consumed_at is None
        assert token.superseded_at is None


# Late import to avoid an unused-import warning at the top of the module.
import dataclasses  # noqa: E402
