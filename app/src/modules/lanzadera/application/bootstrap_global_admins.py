# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin use case: bootstrap_global_admins (D91, DA-11).

Reads ``GLOBAL_ADMIN_EMAILS`` (semicolon-separated) and ensures each
email is a global admin. Idempotent across restarts: existing users
are left alone (their status, password_hash, etc. are not mutated),
and existing global-admin rows are not duplicated. The function
returns the number of NEW global admin rows created (zero on every
re-run after the first).

D91: the CLI ``set-password`` is the exclusive path for provisioning
the first admin. After that, this use case is the way new admins
join the platform — usually triggered from a config-management tool
that pushes the same ``GLOBAL_ADMIN_EMAILS`` to every container.
"""

from __future__ import annotations

import os
from datetime import datetime
from typing import Protocol
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    AuditLogEntry,
    GlobalAdminRepository,
    UserRepository,
)

# SecretManager is replaced with the narrow inline _DniCipher Protocol;
# the real implementation is NationalIdCipher (CA-S2 + D88).
from app.src.modules.lanzadera.domain.user import User, UserStatus


class _DniCipher(Protocol):
    """Narrow port — only the encrypt op the bootstrap needs."""

    def encrypt(self, plaintext: str) -> bytes: ...


def _read_global_admin_emails(env_var: str = "GLOBAL_ADMIN_EMAILS") -> list[str]:
    """Parse ``GLOBAL_ADMIN_EMAILS`` into a normalised list of emails.

    Empty / unset env var returns an empty list — bootstrap is a no-op
    (the spec says "if ``GLOBAL_ADMIN_EMAILS`` becomes unset, the
    adapter returns without error and no user is created").
    """
    raw = os.environ.get(env_var, "")
    if not raw:
        return []
    return [email.strip().lower() for email in raw.split(";") if email.strip()]


async def bootstrap_global_admins(
    *,
    users: UserRepository,
    global_admins: GlobalAdminRepository,
    audit: AuditLog,
    secrets: _DniCipher,
    now: datetime,
    actor_id: UUID | None = None,
    env_var: str = "GLOBAL_ADMIN_EMAILS",
    placeholder_dni: str = "BOOTSTRAP-PLACEHOLDER",
) -> int:
    """Ensure every email in ``env_var`` is a global admin.

    Returns the number of NEW global-admin rows created. Existing
    users are not modified (the spec is explicit: idempotent across
    restarts, no row mutation). Missing users are created with
    ``status = password_reset_required`` and ``password_hash = NULL`` —
    the admin will need to set their password via the CLI before they
    can log in (D91).
    """
    emails = _read_global_admin_emails(env_var)
    created = 0
    for email in emails:
        existing = await users.get_by_email(email)
        if existing is None:
            user = User(
                id=uuid4(),
                email=email,
                name=email.split("@")[0],
                dni_encrypted=secrets.encrypt(placeholder_dni),
                password_hash=None,
                status=UserStatus.PASSWORD_RESET_REQUIRED,
                failed_attempts=0,
                last_login_at=None,
                created_at=now,
                updated_at=now,
            )
            await users.create(user)
            existing = user
        if not await global_admins.is_global_admin(existing.id):
            await global_admins.grant(existing.id)
            created += 1
            await audit.append(
                AuditLogEntry(
                    event_type="global_admins.bootstrap",
                    actor_id=actor_id,
                    target_id=str(existing.id),
                    module="lanzadera",
                    result="success",
                    correlation_id=uuid4(),
                    payload={"email": existing.email, "source": "env"},
                    created_at=now,
                )
            )
    return created


__all__ = ["bootstrap_global_admins"]
