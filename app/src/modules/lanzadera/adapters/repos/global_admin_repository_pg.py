"""PostgreSQL adapter for :class:`GlobalAdminRepositoryPort`.

D21 + D42: global-admin membership is a row in ``global_admins``.
``revoke`` rejects when the operation would leave the system without
any global admin (D42).

**Skeleton** — the bodies raise ``NotImplementedError`` until migration
``0001_core_schema`` (already on main) extends the ``global_admins``
table with the last-admin constraint. The contract is declared today
so the bootstrap CLI can pin its tests against
``typing.Protocol`` structural conformance.
"""

from __future__ import annotations

from collections.abc import Sequence
from uuid import UUID

from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)


class GlobalAdminRepositoryPg(GlobalAdminRepositoryPort):
    """SQLAlchemy Core + asyncpg implementation. Skeleton — see module docstring."""

    async def list_all(self) -> Sequence[GlobalAdmin]:
        raise NotImplementedError(
            "GlobalAdminRepositoryPg.list_all awaits last-admin constraint wiring"
        )

    async def is_global_admin(self, user_id: UUID) -> bool:
        raise NotImplementedError(
            "GlobalAdminRepositoryPg.is_global_admin awaits last-admin constraint wiring"
        )

    async def grant(self, user_id: UUID) -> None:
        raise NotImplementedError(
            "GlobalAdminRepositoryPg.grant awaits last-admin constraint wiring"
        )

    async def revoke(self, user_id: UUID) -> None:
        raise NotImplementedError(
            "GlobalAdminRepositoryPg.revoke awaits last-admin constraint wiring"
        )
