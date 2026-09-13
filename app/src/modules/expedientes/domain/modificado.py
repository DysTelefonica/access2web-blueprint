"""Modificado child entity — D-EXP-1, CAP-009.

Legacy ``TbExpedientesModificados`` carries:

- IDExpedienteModificado (PK)
- IDExpediente (FK, nullable)
- NModificado (255)
- FechaFirmaModificado / FechaFinModificado
- Descripcion (memo)

The legacy schema treats ``NModificado`` and the dates as nullable
(D-EXP-2), so the domain preserves nulls. ``descripcion`` is unbounded
text in Access; we cap it at 1024 chars to keep the read model compact
without dropping valid legacy data.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from uuid import UUID


@dataclass
class Modificado:
    """CAP-009 modification history entry attached to an Expediente."""

    id: UUID
    id_expediente: UUID
    n_modificado: str | None
    fecha_firma: date | None
    fecha_fin: date | None
    descripcion: str | None
    created_at: datetime | None = None
    updated_at: datetime | None = None

    def __post_init__(self) -> None:
        if self.id is None or self.id_expediente is None:
            raise ValueError("id and id_expediente are required")
        self._normalize_n_modificado()
        self._validate_dates()
        self._validate_descripcion()

    def _normalize_n_modificado(self) -> None:
        if self.n_modificado is None:
            return
        cleaned = self.n_modificado.strip()
        if not cleaned:
            raise ValueError("n_modificado must not be blank")
        if len(cleaned) > 255:
            raise ValueError("n_modificado must not exceed 255 characters")
        self.n_modificado = cleaned

    def _validate_dates(self) -> None:
        if self.fecha_firma is None or self.fecha_fin is None:
            return
        if self.fecha_fin < self.fecha_firma:
            raise ValueError("fecha_fin must not be before fecha_firma")

    def _validate_descripcion(self) -> None:
        if self.descripcion is not None and len(self.descripcion) > 1024:
            raise ValueError("descripcion must not exceed 1024 characters")
