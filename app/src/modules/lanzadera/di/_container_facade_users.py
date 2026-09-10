# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W65
"""User-identity use case methods + port accessors (W65 #596).

W65 (#596): cleave LanzaderaContainer by port group.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING, Any, cast

if TYPE_CHECKING:
    from uuid import UUID

    from app.src.modules.lanzadera.domain.ports import UserRepository
    from app.src.modules.lanzadera.domain.user import User

# Runtime imports for ports used in cast() at runtime.
from app.src.modules.lanzadera.domain.ports import AuditLog
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)


class LanzaderaFacadeUsers:
    """User-identity use cases and read-only port accessors.

    ``LanzaderaContainer`` creates one instance per container; callers
    access it via ``container._facade_users`` (private) or implicitly through
    the ``__getattr__`` delegation on ``LanzaderaContainer``.
    """

    def __init__(
        self,
        use_cases: dict[str, Any],
        clock: Any,
        user_repo: Any,
        global_admin_repo: Any,
        audit: Any,
    ) -> None:
        self._use_cases = use_cases
        self._clock = clock
        self._user_repo = user_repo
        self._global_admin_repo = global_admin_repo
        self._audit = audit

    # -- user use cases -------------------------------------------------------

    async def create_user(
        self,
        email: str,
        *,
        name: str,
        national_id: str,
        actor_id: UUID | None = None,
    ) -> User:
        return await self._use_cases["create_user"](  # type: ignore[no-any-return]
            email=email,
            name=name,
            national_id=national_id,
            actor_id=actor_id,
        )

    async def disable_user(
        self,
        user_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["disable_user"](
            user_id=user_id,
            actor_id=actor_id,
            now=self._clock(),
        )

    async def grant_global_admin(
        self,
        user_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["grant_global_admin"](
            user_id=user_id,
            actor_id=actor_id,
            now=self._clock(),
        )

    async def revoke_global_admin(
        self,
        user_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["revoke_global_admin"](
            user_id=user_id,
            actor_id=actor_id,
            now=self._clock(),
        )

    async def list_all_users(
        self,
        limit: int = 50,
        offset: int = 0,
    ) -> tuple[Sequence[User], int]:
        return await self._user_repo.list_all_paginated(limit=limit, offset=offset)  # type: ignore[no-any-return]

    # -- user port accessors --------------------------------------------------

    @property
    def users(self) -> UserRepository:
        return cast(UserRepository, self._user_repo)

    @property
    def global_admin_repo(self) -> GlobalAdminRepositoryPort:
        return cast(GlobalAdminRepositoryPort, self._global_admin_repo)

    @property
    def audit_repo(self) -> AuditLog:
        return cast(AuditLog, self._audit)

    @property
    def audit(self) -> AuditLog:
        return cast(AuditLog, self._audit)
