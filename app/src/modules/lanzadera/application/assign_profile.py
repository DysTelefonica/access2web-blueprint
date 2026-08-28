# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin use case: assign_profile (D22, DA-12, H11, DA-11).

Attaches a (user, app, profile) triple. D22 + H11: the assignment is
the only source of truth for effective permissions. DA-12: SinAcceso
exclusivity is enforced at the application layer — the use case
rejects a profile whose capabilities contain ``SinAcceso`` for a
user who already has any non-SinAcceso capability on the same app.

D5: the row is unique on ``(user_id, app_id)``; the repository
raises on duplicate. The audit row is appended alongside the insert.
"""

from __future__ import annotations

from datetime import datetime
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.errors import (
    ProfileNotFoundError,
    UserNotFoundError,
)
from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    AuditLogEntry,
    UserRepository,
)
from app.src.modules.lanzadera.domain.ports.app_repository import (
    AppRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.profile_repository import (
    ProfileRepositoryPort,
)


async def assign_profile(
    user_id: UUID,
    app_id: int,
    profile_id: UUID,
    *,
    users: UserRepository,
    apps: AppRepositoryPort,
    profiles: ProfileRepositoryPort,
    assignments: AssignmentRepositoryPort,
    audit: AuditLog,
    now: datetime,
    actor_id: UUID | None = None,
) -> None:
    """Assign the profile to the user for the given app.

    DA-12 SinAcceso exclusivity: if the profile is the ``SinAcceso``
    sentinel (capabilities == ``{"SinAcceso": True}``) the user must
    not have any other non-SinAcceso assignment for the same app. We
    reject the assignment with ``ValueError`` before touching the
    database so the audit row can record the rejection.

    Raises ``UserNotFoundError`` / ``ProfileNotFoundError`` for missing
    rows. The repository raises on duplicate ``(user_id, app_id)``.
    """
    user = await users.get_by_id(user_id)
    if user is None:
        raise UserNotFoundError(f"no user with id {user_id!r}")
    app = await apps.get_by_id(app_id)
    if app is None:
        raise ValueError(f"no app with id {app_id!r}")
    profile = await profiles.get_by_id(profile_id)
    if profile is None:
        raise ProfileNotFoundError(f"no profile with id {profile_id!r}")
    if profile.app_id != app_id:
        raise ProfileNotFoundError(
            f"profile {profile_id!r} belongs to app {profile.app_id!r}, not {app_id!r}"
        )

    if profile.capabilities == {"SinAcceso": True}:
        existing = await assignments.list_for_user(user_id)
        for assignment in existing:
            if assignment.app_id != app_id:
                continue
            other_profile = await profiles.get_by_id(assignment.profile_id)
            if other_profile is None:
                continue
            if other_profile.capabilities != {"SinAcceso": True}:
                raise ValueError(
                    f"cannot assign SinAcceso to user {user_id!r} on app {app_id!r}: "
                    f"profile {assignment.profile_id!r} is not SinAcceso"
                )

    await assignments.create(user_id, app_id, profile_id)
    await audit.append(
        AuditLogEntry(
            event_type="assignments.create",
            actor_id=actor_id,
            target_id=f"{user_id}:{app_id}:{profile_id}",
            module="lanzadera",
            result="success",
            correlation_id=uuid4(),
            payload={"user_email": user.email, "app_id": app_id, "profile_id": str(profile_id)},
            created_at=now,
        )
    )


__all__ = ["assign_profile"]
