# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp
"""Admin use case: revoke_assignment (D22, DA-12, DA-11).

Soft-deletes a user's live assignment for a given app by setting
``revoked_at = now``. The row stays in the table for audit trail
(DA-12). The audit row records who performed the revocation and which
(user, app) pair was affected.

Raises ``ValueError`` when no live assignment exists for the pair.
"""

from __future__ import annotations

from datetime import datetime
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    AuditLogEntry,
)
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)


async def revoke_assignment(
    user_id: UUID,
    app_id: int,
    *,
    assignments: AssignmentRepositoryPort,
    audit: AuditLog,
    now: datetime,
    actor_id: UUID | None = None,
) -> None:
    """Revoke the user's live assignment for ``app_id``.

    Idempotent: if the user has no live assignment for the app, the
    function raises ``ValueError`` so the audit row can record the
    rejection. This lets the caller distinguish "already revoked" from
    "never existed" without a separate lookup.

    Args:
        user_id: the user whose assignment to revoke.
        app_id: the app whose assignment to revoke.
        assignments: the assignment repository.
        audit: the audit log.
        now: current timestamp (injected for deterministic tests).
        actor_id: the global admin performing the revocation
            (``request.state.user_id`` in the HTTP layer).

    Raises:
        ValueError: no live assignment exists for ``(user_id, app_id)``.
    """
    assignment = await assignments.revoke(user_id, app_id, now=now)
    if assignment is None:
        raise ValueError(f"no live assignment for user={user_id!r} app={app_id!r}")

    await audit.append(
        AuditLogEntry(
            event_type="assignments.revoke",
            actor_id=actor_id,
            target_id=f"{user_id}:{app_id}",
            module="lanzadera",
            result="success",
            correlation_id=uuid4(),
            payload={"user_id": str(user_id), "app_id": app_id},
            created_at=now,
        )
    )


__all__ = ["revoke_assignment"]
