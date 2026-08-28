# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin use case: list_effective_apps (DA-8).

Returns every app the user can reach. The hot read path goes through
the cache layer (60 s TTL) above this use case; the use case itself
is the application-level entry point. D42: revoked assignments are
filtered out at the repository level (``list_for_user`` only returns
``revoked_at IS NULL`` rows).
"""

from __future__ import annotations

from collections.abc import Sequence
from uuid import UUID

from app.src.modules.lanzadera.domain.app import App
from app.src.modules.lanzadera.domain.ports.app_repository import (
    AppRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)


async def list_effective_apps(
    user_id: UUID,
    *,
    apps: AppRepositoryPort,
    assignments: AssignmentRepositoryPort,
) -> Sequence[App]:
    """Return the live apps the user is assigned to (revoked excluded).

    Pulls the user's non-revoked assignments, then resolves each
    ``(user_id, app_id)`` pair to the matching ``App`` row. Apps the
    user is assigned to but no longer exist in the catalog are
    silently dropped (the FK is enforced at insert time, so this
    is a defensive-only check).
    """
    live = await assignments.list_for_user(user_id)
    result: list[App] = []
    for assignment in live:
        app = await apps.get_by_id(assignment.app_id)
        if app is not None:
            result.append(app)
    return tuple(result)


__all__ = ["list_effective_apps"]
