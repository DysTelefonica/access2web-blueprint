# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W65
"""App-assignment / bootstrap use case methods + port accessors (W65 #596).

W65 (#596): cleave LanzaderaContainer by port group.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING, Any, cast

if TYPE_CHECKING:
    from uuid import UUID

    from app.src.modules.lanzadera.domain.app import App
    from app.src.modules.lanzadera.domain.ports.app_repository import AppRepositoryPort
    from app.src.modules.lanzadera.domain.ports.assignment_repository import (
        AssignmentRepositoryPort,
    )
    from app.src.modules.lanzadera.domain.ports.presence_repository import PresenceRepository

# Runtime import for AuditLog (avoids mypy import-not-found in TYPE_CHECKING).
from app.src.modules.lanzadera.domain.ports import AuditLog


class LanzaderaFacadeApps:
    """App-assignment / bootstrap use cases and read-only port accessors.

    ``LanzaderaContainer`` creates one instance per container; callers
    access it via ``container._facade_apps`` (private) or implicitly through
    the ``__getattr__`` delegation on ``LanzaderaContainer``.
    """

    def __init__(
        self,
        use_cases: dict[str, Any],
        clock: Any,
        app_repo: Any,
        assignment_repo: Any,
        audit: Any,
        presence_repo: Any,
    ) -> None:
        self._use_cases = use_cases
        self._clock = clock
        self._app_repo = app_repo
        self._assignment_repo = assignment_repo
        self._audit = audit
        self._presence_repo = presence_repo

    # -- app-assignment use cases ----------------------------------------------

    async def assign_profile(
        self,
        user_id: UUID,
        app_id: int,
        profile_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["assign_profile"](
            user_id=user_id,
            app_id=app_id,
            profile_id=profile_id,
            actor_id=actor_id,
            now=self._clock(),
        )

    async def list_effective_apps(self, user_id: UUID) -> Sequence[App]:
        return await self._use_cases["list_effective_apps"](user_id=user_id)  # type: ignore[no-any-return]

    async def revoke_assignment(
        self,
        user_id: UUID,
        app_id: int,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["revoke_assignment"](
            user_id=user_id,
            app_id=app_id,
            actor_id=actor_id,
            now=self._clock(),
        )

    # -- bootstrap / audit ---------------------------------------------------

    async def bootstrap_global_admins(
        self,
        *,
        actor_id: UUID | None = None,
    ) -> int:
        return await self._use_cases["bootstrap_global_admins"](  # type: ignore[no-any-return]
            actor_id=actor_id,
            now=self._clock(),
        )

    async def audit_append(
        self,
        event_type: str,
        *,
        actor_id: UUID | None,
        target_id: str,
        result: str,
        payload: dict[str, object],
        correlation_id: UUID | None = None,
        module: str = "lanzadera",
    ) -> AuditLog:
        return await self._use_cases["audit_append"](  # type: ignore[no-any-return]
            event_type=event_type,
            actor_id=actor_id,
            target_id=target_id,
            result=result,
            payload=payload,
            correlation_id=correlation_id,
            module=module,
            now=self._clock(),
        )

    # -- app port accessors ---------------------------------------------------

    @property
    def app_repo(self) -> AppRepositoryPort:
        return cast(AppRepositoryPort, self._app_repo)

    @property
    def assignment_repo(self) -> AssignmentRepositoryPort:
        return cast(AssignmentRepositoryPort, self._assignment_repo)

    @property
    def audit_repo(self) -> AuditLog:
        return cast(AuditLog, self._audit)

    @property
    def presence_repo(self) -> PresenceRepository:
        return cast(PresenceRepository, self._presence_repo)
