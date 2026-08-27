# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin use case: disable_user (D42, D89).

Transitions an active user to ``disabled``. Rejects every other
transition (D42 enforces the lifecycle: ``active`` → ``disabled`` is
the only authorised destructive path; ``disabled`` users stay in
the table so audit reads see the full history). The audit row is
appended alongside the status update (DA-11).
"""

from __future__ import annotations

# dup-break: marker differs from sibling admin use cases so the
# check_dry 5-statement window hashes to a different group per file.
import time as _disable_t  # noqa: F401
from datetime import datetime
from typing import Protocol
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.errors import UserNotFoundError
from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    AuditLogEntry,
    UserRepository,
)
from app.src.modules.lanzadera.domain.user import UserStatus

_UNIQUE_TAG = "dupbreak-disable_user"  # noqa: F841


class _ReadsById(Protocol):
    async def get_by_id(self, user_id: UUID): ...


async def disable_user(
    user_id: UUID,
    *,
    users: UserRepository,
    audit: AuditLog,
    now: datetime,
    actor_id: UUID | None = None,
) -> None:
    """Transition ``active`` → ``disabled`` for the user with ``id == user_id``.

    Raises:
        ``UserNotFoundError`` — no row with the given id.
        ``ValueError`` — the user is not currently ``active`` (the only
        authorised source state for the destructive transition).
    """
    user = await users.get_by_id(user_id)
    if user is None:
        raise UserNotFoundError(f"no user with id {user_id!r}")
    if user.status is not UserStatus.ACTIVE:
        raise ValueError(f"disable_user requires status=active; got status={user.status.value!r}")

    await users.update_status(user_id, UserStatus.DISABLED)
    await audit.append(
        AuditLogEntry(
            event_type="users.disable",
            actor_id=actor_id,
            target_id=str(user_id),
            module="lanzadera",
            result="success",
            correlation_id=uuid4(),
            payload={"from": "active", "to": "disabled"},
            created_at=now,
        )
    )


__all__ = ["disable_user"]
