"""Expediente aggregate root — D-EXP-1, D-EXP-4, CAP-001..CAP-007.

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID

from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


@dataclass
class Expediente:
    """Aggregate root.

    Mutable: estado transitions and version bump happen at the application layer.
    Version >= 1 (D-EXP-4: optimistic locking).
    Hierarchy (CAP-006): LOTE and BASED require id_expediente_padre; AM does not.
    """

    id: UUID
    tipo: ExpedienteTipo
    estado: ExpedienteEstado
    version: int
    id_expediente_padre: UUID | None = None
    # Timestamps
    created_at: datetime | None = None
    updated_at: datetime | None = None

    def __post_init__(self) -> None:
        if self.version < 1:
            raise ValueError("version must be >= 1")
        if self.tipo.value in ("LOTE", "BASED") and self.id_expediente_padre is None:
            raise ValueError(f"tipo={self.tipo.value} requires id_expediente_padre")
