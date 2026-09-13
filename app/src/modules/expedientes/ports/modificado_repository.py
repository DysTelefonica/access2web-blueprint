"""ModificadoRepositoryPort — CAP-009, D-EXP-1."""

from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from uuid import UUID


class ModificadoRepositoryPort(Protocol):
    """CAP-009 modification history persistence port."""

    async def get_by_id(self, modificado_id: UUID) -> object | None: ...
    async def get_by_expediente(self, expediente_id: UUID) -> list[object]: ...
    async def upsert(self, modificado: object) -> object: ...
