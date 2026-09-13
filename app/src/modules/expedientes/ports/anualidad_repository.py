"""AnualidadRepositoryPort — CAP-011, D-EXP-1."""

from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from uuid import UUID


class AnualidadRepositoryPort(Protocol):
    """CAP-011 anualidad persistence port."""

    async def get_by_id(self, anualidad_id: UUID) -> object | None: ...
    async def get_by_expediente(self, expediente_id: UUID) -> list[object]: ...
    async def upsert(self, anualidad: object) -> object: ...
