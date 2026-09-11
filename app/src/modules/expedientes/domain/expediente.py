"""Expediente aggregate root — D-EXP-1, D-EXP-4, CAP-001..CAP-008.

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from uuid import UUID

from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo
from app.src.modules.expedientes.domain.hito import Hito


@dataclass
class Expediente:
    """Aggregate root.

    Mutable: estado transitions and version bump happen at the application layer.
    Version >= 1 (D-EXP-4: optimistic locking).
    Hierarchy (CAP-006): LOTE and BASED require id_expediente_padre; AM does not.
    Hitos (CAP-008): milestones attached to this Expediente, added via add_hito().
    """

    id: UUID
    tipo: ExpedienteTipo
    estado: ExpedienteEstado
    version: int
    id_expediente_padre: UUID | None = None
    created_at: datetime | None = None
    updated_at: datetime | None = None
    _hitos: list[Hito] = field(default_factory=list)

    @property
    def hitos(self) -> tuple[Hito, ...]:
        """Immutable snapshot of hitos attached to this Expediente."""
        return tuple(self._hitos)

    def add_hito(self, hito: Hito) -> None:
        """Attach a Hito to this Expediente (CAP-008).

        Raises ValueError if hito.id_expediente does not match self.id
        or if a Hito with the same id is already attached.
        """
        if hito.id_expediente != self.id:
            raise ValueError(
                f"wrong Expediente {hito.id_expediente} for Hito {hito.id}; expected {self.id}"
            )
        if any(h.id == hito.id for h in self._hitos):
            raise ValueError(f"Hito {hito.id} already exists in Expediente {self.id}")
        self._hitos.append(hito)

    def remove_hito(self, hito_id: UUID) -> None:
        """Detach a Hito from this Expediente by its id.

        Raises ValueError if no Hito with that id is attached.
        """
        for i, h in enumerate(self._hitos):
            if h.id == hito_id:
                self._hitos.pop(i)
                return
        raise ValueError(f"Hito {hito_id} not found in Expediente {self.id}")

    def __post_init__(self) -> None:
        if self.version < 1:
            raise ValueError("version must be >= 1")
        if self.tipo.value in ("LOTE", "BASED") and self.id_expediente_padre is None:
            raise ValueError(f"tipo={self.tipo.value} requires id_expediente_padre")
