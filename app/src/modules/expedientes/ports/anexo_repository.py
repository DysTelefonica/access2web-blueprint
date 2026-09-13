"""AnexoRepositoryPort — CAP-010, D-EXP-1."""

from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from uuid import UUID


class AnexoRepositoryPort(Protocol):
    """CAP-010 anexo persistence port."""

    async def get_by_id(self, anexo_id: UUID) -> object | None: ...
    async def get_by_expediente(self, expediente_id: UUID) -> list[object]: ...
    async def add(self, anexo: object) -> object: ...
    async def delete(self, anexo_id: UUID) -> None: ...
