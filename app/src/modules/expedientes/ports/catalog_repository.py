"""CatalogRepositoryPort — CAP-015..CAP-024, D-EXP-1."""

from __future__ import annotations

from typing import Protocol


class CatalogRepositoryPort(Protocol):
    """Puerto genérico para las 14 tablas maestras. Valor-inmutable con vigencia temporal."""

    async def list_all(self) -> list[object]: ...
    async def get_by_id(self, catalog_id: int) -> object | None: ...
    async def search(self, query: str, limit: int = 20) -> list[object]: ...
