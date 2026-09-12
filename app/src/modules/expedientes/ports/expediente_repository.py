"""ExpedienteRepositoryPort — D-EXP-1, D-EXP-4, CAP-001..CAP-007."""

from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from uuid import UUID


class ExpedienteRepositoryPort(Protocol):
    """Agregado Expediente: cabecera, jerarquía, entidades y hijos. CRUD atómico (D-EXP-4)."""

    async def get_by_id(self, expediente_id: UUID) -> object | None: ...
    async def create(self, aggregate: object) -> object: ...
    async def update(self, aggregate: object) -> object: ...
    async def delete(self, expediente_id: UUID) -> None: ...
    async def list_by_state(
        self, estado: str, limit: int, offset: int
    ) -> tuple[list[object], int]: ...
