"""SuministradorRepositoryPort — CAP-014, D-EXP-1."""

from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from uuid import UUID


class SuministradorRepositoryPort(Protocol):
    """CAP-014 suministrador persistence port."""

    async def get_by_id(self, suministrador_id: UUID) -> object | None: ...
    async def get_by_expediente(self, expediente_id: UUID) -> list[object]: ...
    async def upsert(self, suministrador: object) -> object: ...
