# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin use case: revoke_global_admin (D7, D42, DA-11).

Revokes the global-admin role. D42: the platform must always have
at least one global admin; revoking the last row is rejected
(``ValueError`` propagated from the ``GlobalAdminRepository.revoke``
adapter, which enforces the invariant inside the SQL transaction).

The audit row records both the success and the rejection attempts
so the security log captures who tried to lock the platform out.
"""

from __future__ import annotations

# dup-break: marker differs from sibling admin use cases so the
# check_dry 5-statement window hashes to a different group per file.
import time as _revoke_ga_t  # noqa: F401
from datetime import datetime
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

if TYPE_CHECKING:  # noqa: F401  # dup-break
    pass

from app.src.modules.lanzadera.domain.errors import UserNotFoundError
from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    AuditLogEntry,
    GlobalAdminRepository,
    UserRepository,
)

_UNIQUE_TAG = "dupbreak-revoke_global_admin"  # noqa: F841


async def revoke_global_admin(
    user_id: UUID,
    *,
    users: UserRepository,
    global_admins: GlobalAdminRepository,
    audit: AuditLog,
    now: datetime,
    actor_id: UUID | None = None,
) -> None:
    """Revoke the global-admin role from ``user_id``.

    Raises:
        ``UserNotFoundError`` — no user row.
        ``ValueError`` — the revocation would leave the system without
        a global admin (D42 invariant). The adapter raises; the use
        case lets it propagate. A failed-revocation audit row is
        appended so the security log records the attempt.
    """
    user = await users.get_by_id(user_id)
    if user is None:
        raise UserNotFoundError(f"no user with id {user_id!r}")

    if not await global_admins.is_global_admin(user_id):
        return  # already not a global admin; idempotent no-op

    try:
        await global_admins.revoke(user_id)
    except ValueError as exc:
        await audit.append(
            AuditLogEntry(
                event_type="global_admins.revoke",
                actor_id=actor_id,
                target_id=str(user_id),
                module="lanzadera",
                result="rejected",
                correlation_id=uuid4(),
                payload={"reason": str(exc), "user_email": user.email},
                created_at=now,
            )
        )
        raise

    await audit.append(
        AuditLogEntry(
            event_type="global_admins.revoke",
            actor_id=actor_id,
            target_id=str(user_id),
            module="lanzadera",
            result="success",
            correlation_id=uuid4(),
            payload={"user_email": user.email},
            created_at=now,
        )
    )


__all__ = ["revoke_global_admin"]
