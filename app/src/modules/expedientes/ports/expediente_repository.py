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
    async def search(
        self,
        *,
        estado: str | None,
        codigo: str | None,
        responsable_id: object | None,
        juridica_id: object | None,
        suministrador_id: object | None,
        limit: int,
        offset: int,
    ) -> tuple[list[object], int]: ...
    async def has_children(self, expediente_id: UUID) -> bool:
        """Return True if the expediente has children of any kind.

        Children include: anexos (TbExpedientesAnexos), hitos
        (TbExpedientesHitos), hijos del agregado (LOTE/BASED with
        ``id_expediente_padre`` pointing here), and any related-data
        vertical that lives under the expediente.

        CAP-003 §Camino feliz requires "impedir pérdida de hijos".
        The use case calls this method before delete; the caller can
        override the check by passing ``force=True`` to the command.
        """
        ...
