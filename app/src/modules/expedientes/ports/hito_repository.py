"""HitoRepositoryPort — CAP-008, D-EXP-1."""

from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from uuid import UUID


class HitoRepositoryPort(Protocol):
    """Hijo del agregado Expediente: fechas hito (Inicio, Adjudicación, Formalización, etc.)."""

    async def get_by_expediente(self, expediente_id: UUID) -> list[object]: ...
    async def upsert(self, hito: object) -> object: ...
    async def delete(self, hito_id: UUID) -> None: ...
