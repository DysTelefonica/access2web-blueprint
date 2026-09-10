# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""User-facing use cases: resolve own apps and capabilities (D22, DA-12, H11).

``get_my_apps`` returns every app the authenticated user can access, with
their profile and capability set. The caller is the HTTP delivery layer
(``GET /auth/me/apps``).

``get_my_capabilities_for_app`` returns the capability names for one app
only. The caller is ``GET /auth/me/apps/{app_id}/capabilities`` and
future app-consumers that need to gate a single resource.

D22 + H11: the assignment is the only source of truth for effective
permissions; ``SinAcceso`` assignments collapse to
``capabilities=["SinAcceso"]`` (DA-12 exclusive cortocircuito).
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from uuid import UUID

from app.src.modules.lanzadera.domain.ports.app_repository import (
    AppRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.profile_repository import (
    ProfileRepositoryPort,
)


@dataclass(frozen=True)
class EffectiveApp:
    """One app visible to the authenticated user."""

    app_id: int
    app_name: str
    app_short_code: str
    profile_code: str
    profile_name: str
    capabilities: tuple[str, ...]


async def get_my_apps(
    user_id: UUID,
    *,
    assignments: AssignmentRepositoryPort,
    apps: AppRepositoryPort,
    profiles: ProfileRepositoryPort,
) -> Sequence[EffectiveApp]:
    """Return the live apps the user is assigned to, with profile and capabilities.

    Iterates over the user's non-revoked assignments, resolves the app and
    profile for each, and returns the effective permission triple. Apps
    whose catalog row has been retired are silently dropped; deactivated
    profiles are also filtered (the ``effective_permissions`` SQL in the
    Postgres adapter already filters ``active=True``).
    """
    live = await assignments.list_for_user(user_id)
    result: list[EffectiveApp] = []
    for assignment in live:
        app = await apps.get_by_id(assignment.app_id)
        if app is None:
            continue
        profile = await profiles.get_by_id(assignment.profile_id)
        if profile is None:
            continue
        if not profile.active:
            continue
        result.append(
            EffectiveApp(
                app_id=app.id,
                app_name=app.name,
                app_short_code=app.short_code,
                profile_code=profile.code,
                profile_name=profile.name,
                capabilities=tuple(sorted(profile.capabilities.keys())),
            )
        )
    return tuple(result)


async def get_my_capabilities_for_app(
    user_id: UUID,
    app_id: int,
    *,
    assignments: AssignmentRepositoryPort,
) -> Sequence[str]:
    """Return the capability names for ``user_id`` on ``app_id``.

    Delegates to ``AssignmentRepositoryPort.effective_permissions`` which
    joins assignments → profiles and flattens the JSONB capabilities keys.
    Returns an empty tuple when the user has no live assignment for the app.
    """
    return await assignments.effective_permissions(user_id, app_id)


__all__ = ["get_my_apps", "get_my_capabilities_for_app", "EffectiveApp"]
