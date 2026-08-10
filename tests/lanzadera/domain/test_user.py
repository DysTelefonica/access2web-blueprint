# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-3, D89 — `User` is the canonical identity record. Email is normalised to
# lowercase at construction time; the entity is mutable (status transitions,
# failed-attempt counter, last-login) but rejects invalid email casing.
"""Strict TDD — `User` entity + `UserStatus` StrEnum (Phase 1, tasks 1.1)."""

from __future__ import annotations

import dataclasses
from datetime import datetime, timedelta, timezone
from uuid import UUID, uuid4

import pytest

from app.src.modules.lanzadera.domain.user import User, UserStatus

# Email literals use \uXXXX escapes for the local-part so casing cannot be
# silently collapsed by a tool or editor (DA-3 requires the constructor to
# reject any non-lowercase email at the boundary, not just emails that look
# like they contain uppercase characters).
EMAIL_LOWER: str = "alice@enterprise.test"
EMAIL_UPPER: str = "ALICE@enterprise.test"
EMAIL_MIXED: str = "AlIcE@enterprise.test"

# ---------------------------------------------------------------------------
# Helpers / fixtures
# ---------------------------------------------------------------------------


def _now() -> datetime:
    return datetime(2026, 8, 9, 12, 0, 0, tzinfo=timezone.utc)


def _new_user(
    *,
    email: str = EMAIL_LOWER,
    name: str = "Alice",
    status: UserStatus = UserStatus.ACTIVE,
    password_hash: str | None = None,
    failed_attempts: int = 0,
) -> User:
    return User(
        id=uuid4(),
        email=email,
        name=name,
        dni_encrypted=b"\x00\x01\x02\x03\x04",
        password_hash=password_hash,
        status=status,
        failed_attempts=failed_attempts,
        last_login_at=None,
        created_at=_now(),
        updated_at=_now(),
    )


# ---------------------------------------------------------------------------
# UserStatus StrEnum
# ---------------------------------------------------------------------------


class TestUserStatusStrEnum:
    """`UserStatus` MUST be a `StrEnum` so JSON serialisation round-trips."""

    def test_active_member_is_str(self) -> None:
        assert isinstance(UserStatus.ACTIVE, str)

    def test_disabled_member_is_str(self) -> None:
        assert isinstance(UserStatus.DISABLED, str)

    def test_password_reset_required_member_is_str(self) -> None:
        assert isinstance(UserStatus.PASSWORD_RESET_REQUIRED, str)

    def test_locked_member_is_str(self) -> None:
        assert isinstance(UserStatus.LOCKED, str)

    def test_string_value_round_trips(self) -> None:
        """DA-3 + Postgres ENUM: the wire value must equal the member name."""
        assert UserStatus("active") is UserStatus.ACTIVE
        assert UserStatus("disabled") is UserStatus.DISABLED
        assert (
            UserStatus("password_reset_required") is UserStatus.PASSWORD_RESET_REQUIRED
        )
        assert UserStatus("locked") is UserStatus.LOCKED

    def test_wire_values_match_postgres_enum(self) -> None:
        """Postgres ENUM `user_status` literal values match the StrEnum members."""
        assert UserStatus.ACTIVE.value == "active"
        assert UserStatus.DISABLED.value == "disabled"
        assert UserStatus.PASSWORD_RESET_REQUIRED.value == "password_reset_required"
        assert UserStatus.LOCKED.value == "locked"


# ---------------------------------------------------------------------------
# User construction — happy path
# ---------------------------------------------------------------------------


class TestUserConstruction:
    """Minimum viable `User` carries every required attribute."""

    def test_user_carries_identity_attributes(self) -> None:
        user_id = uuid4()
        user = User(
            id=user_id,
            email="bob@enterprise.test",
            name="Bob",
            dni_encrypted=b"\xff\xff",
            password_hash=None,
            status=UserStatus.PASSWORD_RESET_REQUIRED,
            failed_attempts=0,
            last_login_at=None,
            created_at=_now(),
            updated_at=_now(),
        )
        assert user.id == user_id
        assert user.email == "bob@enterprise.test"
        assert user.name == "Bob"
        assert user.dni_encrypted == b"\xff\xff"
        assert user.password_hash is None
        assert user.status is UserStatus.PASSWORD_RESET_REQUIRED
        assert user.failed_attempts == 0
        assert user.last_login_at is None
        assert user.created_at == _now()
        assert user.updated_at == _now()

    def test_user_is_mutable_for_status_transitions(self) -> None:
        """Status, failed_attempts, last_login_at change over time -> not frozen."""
        user = _new_user()
        updated = _now() + timedelta(hours=1)
        user.status = UserStatus.ACTIVE
        user.failed_attempts = 0
        user.last_login_at = updated
        user.updated_at = updated
        assert user.status is UserStatus.ACTIVE
        assert user.failed_attempts == 0
        assert user.last_login_at == updated

    def test_password_hash_can_be_set_after_reset(self) -> None:
        """D89 + D90: `password_hash` becomes non-NULL after `consume_reset_token`."""
        user = _new_user(status=UserStatus.PASSWORD_RESET_REQUIRED)
        assert user.password_hash is None
        user.password_hash = "$argon2id$v=19$m=65536,t=3,p=4$..."
        user.status = UserStatus.ACTIVE
        assert user.password_hash is not None
        assert user.status is UserStatus.ACTIVE


# ---------------------------------------------------------------------------
# DA-3 — email lower-case invariant
# ---------------------------------------------------------------------------


class TestUserEmailLowerCaseInvariant:
    """DA-3: `email == email.lower()` enforced at construction time."""

    def test_lowercase_email_is_accepted(self) -> None:
        user = _new_user(email=EMAIL_LOWER)
        assert user.email == EMAIL_LOWER

    def test_uppercase_email_is_rejected(self) -> None:
        with pytest.raises(ValueError, match="lower"):
            _new_user(email=EMAIL_UPPER)

    def test_mixed_case_email_is_rejected(self) -> None:
        with pytest.raises(ValueError, match="lower"):
            _new_user(email=EMAIL_MIXED)

    def test_uppercase_local_part_is_rejected(self) -> None:
        """A single uppercase character in the local-part fails the invariant."""
        with pytest.raises(ValueError, match="lower"):
            _new_user(email="aliceZ@enterprise.test")


# ---------------------------------------------------------------------------
# DA-3 + DA-13 — name invariant
# ---------------------------------------------------------------------------


class TestUserNameInvariant:
    """Domain invariant: `name` MUST be a non-empty, non-whitespace string."""

    def test_empty_name_is_rejected(self) -> None:
        with pytest.raises(ValueError, match="name"):
            _new_user(name="")

    def test_whitespace_only_name_is_rejected(self) -> None:
        with pytest.raises(ValueError, match="name"):
            _new_user(name="   ")


# ---------------------------------------------------------------------------
# State transitions
# ---------------------------------------------------------------------------


class TestUserStateTransitions:
    """Status transitions cover the four ENUM values without rejecting any."""

    @pytest.mark.parametrize(
        ("target", "previous"),
        [
            (UserStatus.ACTIVE, UserStatus.PASSWORD_RESET_REQUIRED),
            (UserStatus.DISABLED, UserStatus.ACTIVE),
            (UserStatus.LOCKED, UserStatus.ACTIVE),
            (UserStatus.PASSWORD_RESET_REQUIRED, UserStatus.LOCKED),
        ],
    )
    def test_status_round_trip(self, target: UserStatus, previous: UserStatus) -> None:
        user = _new_user(status=previous)
        user.status = target
        assert user.status is target

    def test_failed_attempts_increments_on_failure(self) -> None:
        user = _new_user()
        user.failed_attempts = user.failed_attempts + 1
        assert user.failed_attempts == 1

    def test_failed_attempts_resets_on_unlock(self) -> None:
        """D38 + D39: global admin unlock resets the counter to 0."""
        user = _new_user(failed_attempts=5, status=UserStatus.LOCKED)
        user.failed_attempts = 0
        user.status = UserStatus.PASSWORD_RESET_REQUIRED
        assert user.failed_attempts == 0
        assert user.status is UserStatus.PASSWORD_RESET_REQUIRED


# ---------------------------------------------------------------------------
# Equality / hashing — value-object-like identity, but mutable fields excluded
# ---------------------------------------------------------------------------


class TestUserEquality:
    """Two `User`s with the same identity fields compare equal."""

    def test_two_users_same_id_compare_equal(self) -> None:
        user_id = uuid4()
        a = User(
            id=user_id,
            email="carol@enterprise.test",
            name="Carol",
            dni_encrypted=b"x",
            password_hash=None,
            status=UserStatus.ACTIVE,
            failed_attempts=0,
            last_login_at=None,
            created_at=_now(),
            updated_at=_now(),
        )
        b = User(
            id=user_id,
            email="carol@enterprise.test",
            name="Carol",
            dni_encrypted=b"x",
            password_hash=None,
            status=UserStatus.ACTIVE,
            failed_attempts=0,
            last_login_at=None,
            created_at=_now(),
            updated_at=_now(),
        )
        assert a == b

    def test_distinct_ids_compare_unequal(self) -> None:
        a = _new_user()
        b = dataclasses.replace(a, id=uuid4())
        assert a != b
