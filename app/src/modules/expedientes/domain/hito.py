"""Hito child entity — D-EXP-1, CAP-008.

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from enum import StrEnum
from uuid import UUID

from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado


class HitoEstado(StrEnum):
    """Hito completion state. Values match the Postgres ENUM literals (CAP-008)."""

    PENDIENTE = "PENDIENTE"
    CUMPLIDO = "CUMPLIDO"
    VENCIDO = "VENCIDO"


@dataclass
class Hito:
    """Milestone child entity attached to an Expediente.

    fecha_hito: mandatory date (no time component — CAP-008).
    garantia_fecha_fin: optional; must be strictly after fecha_hito when present.
    """

    id: UUID
    id_expediente: UUID
    fecha_hito: date
    garantia_fecha_fin: date | None
    estado: ExpedienteEstado
    created_at: datetime | None = None
    updated_at: datetime | None = None

    def __post_init__(self) -> None:
        if isinstance(self.fecha_hito, datetime):
            raise ValueError("fecha_hito must be a date (no time component)")
        if not isinstance(self.fecha_hito, date):
            raise ValueError("fecha_hito must be a date")
        if self.garantia_fecha_fin is not None:
            if self.garantia_fecha_fin <= self.fecha_hito:
                raise ValueError(
                    "garantia_fecha_fin must be strictly after fecha_hito"
                )
