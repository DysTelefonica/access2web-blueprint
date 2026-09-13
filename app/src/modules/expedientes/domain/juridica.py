"""Juridica child entity — D-EXP-1, CAP-013.

Legacy ``TbExpedientesJuridicas`` carries:

- IDExpedienteJuridica (PK)
- IDExpediente (FK)
- IDJuridica (FK)
- IDSuministrador (FK, nullable)
- ContratistaPrincipal / SubContratista (Text(2), nullable, Yes/No)

The D102 cross-cutting rule will standardise these to BOOLEAN in
PostgreSQL, but the domain keeps them nullable today. ``id_juridica``
is mandatory because every juridica row must reference a juridica.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID


@dataclass
class Juridica:
    """CAP-013 juridica row attached to an Expediente."""

    id: UUID
    id_expediente: UUID
    id_juridica: UUID
    id_suministrador: UUID | None
    contratista_principal: bool | None
    sub_contratista: bool | None
    created_at: datetime | None = None
    updated_at: datetime | None = None

    def __post_init__(self) -> None:
        if self.id is None or self.id_expediente is None or self.id_juridica is None:
            raise ValueError("id, id_expediente and id_juridica are required")
