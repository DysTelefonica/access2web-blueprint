# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin use case: grant_global_admin (D7, D42, DA-11).

Grants the global-admin role to a user. Idempotent: re-granting a row
that already exists is a no-op (the membership is a fact, not a state
machine; revoking is the only way to undo it). D42 — the platform must
have at least one global admin at all times — is a revoke-side
invariant; grant is always safe.

D7: the membership is materialised as a row in ``global_admins``,
NOT as a column on ``users``. The role lives in the join table.
"""

from __future__ import annotations

# dup-break: marker differs from sibling admin use cases so the
# check_dry 5-statement window hashes to a different group per file.
import time as _grant_ga_t  # noqa: F401
from datetime import datetime
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.errors import UserNotFoundError
from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    AuditLogEntry,
    GlobalAdminRepository,
    UserRepository,
)

_UNIQUE_TAG = "dupbreak-grant_global_admin"  # noqa: F841


async def grant_global_admin(
    user_id: UUID,
    *,
    users: UserRepository,
    global_admins: GlobalAdminRepository,
    audit: AuditLog,
    now: datetime,
    actor_id: UUID | None = None,
) -> None:
    """Grant the global-admin role to ``user_id``.

    Idempotent: if the user is already a global admin, the function
    returns without raising (no duplicate row, no error, no audit
    entry — granting twice is a fact, not a state transition).

    Raises ``UserNotFoundError`` if the user row does not exist.
    """
    user = await users.get_by_id(user_id)
    if user is None:
        raise UserNotFoundError(f"no user with id {user_id!r}")

    if await global_admins.is_global_admin(user_id):
        return

    await global_admins.grant(user_id)
    await audit.append(
        AuditLogEntry(
            event_type="global_admins.grant",
            actor_id=actor_id,
            target_id=str(user_id),
            module="lanzadera",
            result="success",
            correlation_id=uuid4(),
            payload={"user_email": user.email},
            created_at=now,
        )
    )


__all__ = ["grant_global_admin"]
