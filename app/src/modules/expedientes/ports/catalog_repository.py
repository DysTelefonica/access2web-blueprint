"""CatalogRepositoryPort — CAP-015..CAP-024, D-EXP-1."""

from __future__ import annotations

from datetime import date
from typing import Protocol


class CatalogRepositoryPort(Protocol):
    """Puerto genérico para las 14 tablas maestras. Valor-inmutable con vigencia temporal.

    Q01 (CAP-015..019) uses ``list_all`` / ``get_by_id`` / ``search``.
    Q02 (CAP-020..024) adds the temporal validity dimension: every row
    carries an optional ``vigente_desde`` / ``vigente_hasta`` pair that the
    delivery layer filters on. The ``vigentes_en`` method returns only
    rows that are valid at the supplied date. Adapters that do not
    support vigencia (Q01 only) implement the new method as a
    pass-through that returns every row.
    """

    async def list_all(self) -> list[object]: ...
    async def get_by_id(self, catalog_id: int) -> object | None: ...
    async def search(self, query: str, limit: int = 20) -> list[object]: ...
    async def list_vigentes_en(self, fecha: date) -> list[object]:
        """Return rows whose ``vigente_desde`` / ``vigente_hasta`` include ``fecha``."""
        ...

    async def vigentes_search(self, query: str, fecha: date, limit: int = 20) -> list[object]:
        """Combine text filter and vigencia range."""
        ...
