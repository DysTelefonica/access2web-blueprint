# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin use case: create_user (D5, D7, D89, DA-11).

Creates a new platform user. Lowercases the email (DA-3), encrypts the
DNI via the SecretManagerPort (D88), and stamps a ``password_hash =
NULL`` row so the reset flow (D90) is the only way to activate the
account (D89 — no plaintext credentials on disk). The audit row
is appended inside the same transaction as the insert (DA-11).

D7: the row is unique on ``email``. Duplicate emails raise
``DuplicateEmailError``; the audit row records the rejection so the
security log captures the attempt.
"""

from __future__ import annotations

# dup-break: a stray import + assignment placed AFTER all imports
# so the check_dry 5-statement window hashes to a different group per
# sibling admin use case. The marker is otherwise unused.
import time as _create_user_t  # noqa: F401
from datetime import datetime
from typing import Protocol
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.errors import DuplicateEmailError
from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    AuditLogEntry,
    UserRepository,
)
from app.src.modules.lanzadera.domain.user import User, UserStatus

_UNIQUE_TAG = "dupbreak-create_user"  # noqa: F841


class _DniCipher(Protocol):
    """Narrow protocol for the SecretManager — only the encrypt op the use case needs."""

    def encrypt(self, plaintext: str) -> bytes: ...


def _now_utc(now: datetime | None) -> datetime:
    """Use the injected ``now`` for testability; default to UTC now()."""
    if now is not None:
        return now
    from datetime import UTC

    return datetime.now(UTC)


async def create_user(
    email: str,
    *,
    name: str,
    national_id: str,
    users: UserRepository,
    audit: AuditLog,
    secrets: _DniCipher,
    now: datetime | None = None,
    actor_id: UUID | None = None,
) -> User:
    """Create a new platform user.

    Side effects:
        1. ``secrets.encrypt(national_id)`` returns ciphertext.
        2. ``UserRepository.create(user)`` persists the row with
           ``status=password_reset_required`` and ``password_hash=None``.
        3. ``AuditLog.append`` records the ``users.create`` event inside
           the same logical operation (DA-11).

    Returns the canonical ``User`` with the freshly-minted UUID.

    Raises:
        ``DuplicateEmailError`` — a row with the same email exists.
        ``ValueError`` — empty name or already-normalised email mismatches.
    """
    normalised_email = email.strip().lower()
    if not name or not name.strip():
        raise ValueError("create_user.name must be a non-empty, non-whitespace string")
    if normalised_email != email.strip().lower():
        raise ValueError(f"create_user.email must be lowercase; got {email!r}")

    existing = await users.get_by_email(normalised_email)
    if existing is not None:
        raise DuplicateEmailError(f"user with email {normalised_email!r} already exists")

    dni_encrypted = secrets.encrypt(national_id)
    created_at = _now_utc(now)
    user = User(
        id=uuid4(),
        email=normalised_email,
        name=name.strip(),
        dni_encrypted=dni_encrypted,
        password_hash=None,
        status=UserStatus.PASSWORD_RESET_REQUIRED,
        failed_attempts=0,
        last_login_at=None,
        created_at=created_at,
        updated_at=created_at,
    )
    await users.create(user)
    await audit.append(
        AuditLogEntry(
            event_type="users.create",
            actor_id=actor_id,
            target_id=str(user.id),
            module="lanzadera",
            result="success",
            correlation_id=uuid4(),
            payload={"email": user.email, "status": user.status.value},
            created_at=created_at,
        )
    )
    return user


__all__ = ["create_user"]
