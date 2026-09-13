"""ResponsableRepositoryPort — CAP-012, D-EXP-1."""

from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from uuid import UUID


class ResponsableRepositoryPort(Protocol):
    """CAP-012 responsable persistence port."""

    async def get_by_id(self, responsable_id: UUID) -> object | None: ...
    async def get_by_expediente(self, expediente_id: UUID) -> list[object]: ...
    async def upsert(self, responsable: object) -> object: ...
