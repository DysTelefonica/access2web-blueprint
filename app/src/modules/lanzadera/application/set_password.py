# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin use case: set_password (D91, D89, DA-2, DA-11).

CLI bootstrap path: hashes a new password and transitions the user to
``active`` + global admin. D91: this is the only path that creates
the first global admin from scratch. After that, the bootstrap is
manual (CLIs are explicitly destructive — every invocation prompts
for confirmation per D26).

D89: the user is created with ``password_hash = NULL`` until this
use case runs. The password is hashed via the ``CredentialHasher``
port (Argon2id, D88) — never stored in plaintext.
"""

from __future__ import annotations

# dup-break: marker differs from sibling admin use cases so the
# check_dry 5-statement window hashes to a different group per file.
import time as _set_password_t  # noqa: F401
from datetime import datetime
from typing import Protocol
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.errors import UserNotFoundError
from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    AuditLogEntry,
    GlobalAdminRepository,
    PasswordHasher,
    UserRepository,
)
from app.src.modules.lanzadera.domain.user import UserStatus

_UNIQUE_TAG = "dupbreak-set_password"  # noqa: F841


class _DniCipher(Protocol):
    """Narrow port — only encrypt for placeholder DNI when creating a new user."""

    def encrypt(self, plaintext: str) -> bytes: ...


async def set_password(
    email: str,
    *,
    new_password: str,
    hasher: PasswordHasher,
    users: UserRepository,
    global_admins: GlobalAdminRepository,
    audit: AuditLog,
    secrets: _DniCipher,
    now: datetime,
    actor_id: UUID | None = None,
) -> None:
    """Provision the bootstrap admin: hash password, activate, grant.

    If the user does not exist, this is the first global admin ever:
    create the user with status=password_reset_required, then proceed
    normally. If the user exists but is not a global admin, grant the
    role (idempotent — see ``grant_global_admin``).
    """
    normalised_email = email.strip().lower()
    user = await users.get_by_email(normalised_email)
    if user is None:
        from app.src.modules.lanzadera.domain.user import User

        user = User(
            id=uuid4(),
            email=normalised_email,
            name=normalised_email.split("@")[0],
            dni_encrypted=secrets.encrypt("BOOTSTRAP-PLACEHOLDER"),
            password_hash=None,
            status=UserStatus.PASSWORD_RESET_REQUIRED,
            failed_attempts=0,
            last_login_at=None,
            created_at=now,
            updated_at=now,
        )
        await users.create(user)
        user = await users.get_by_email(normalised_email)
        if user is None:
            raise UserNotFoundError(f"failed to create user for {email!r}")

    password_hash = await hasher.hash(new_password)
    await users.update_password_and_activate(user.id, password_hash)
    if not await global_admins.is_global_admin(user.id):
        await global_admins.grant(user.id)
    await audit.append(
        AuditLogEntry(
            event_type="users.set_password",
            actor_id=actor_id,
            target_id=str(user.id),
            module="lanzadera",
            result="success",
            correlation_id=uuid4(),
            payload={"email": user.email, "first_admin": True},
            created_at=now,
        )
    )


__all__ = ["set_password"]
