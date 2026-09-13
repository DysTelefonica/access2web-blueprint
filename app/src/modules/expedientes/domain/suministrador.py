"""Suministrador child entity — D-EXP-1, CAP-014.

Legacy ``TbExpedientesSuministradores`` carries:

- IDExpedienteSuministrador (PK)
- IDExpediente (FK, NOT NULL)
- IDSuministrador (FK, NOT NULL)
- IDPadre (FK, nullable — tree parent)
- Descripcion (legacy typo: ``Descripcon``)
- ContratistaPrincipal / SubContratista (Text(2), nullable, Yes/No)

CAP-014 rule «el árbol manda»: a root row per expediente must mark
``contratista_principal=True``; subcontratistas carry a parent id.
The legacy typo ``Descripcon`` is normalised to ``descripcion`` by
the migration pipeline (per the field-mapping notes); the domain
keeps the canonical name.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID

MAX_DESCRIPCION_LEN = 1024


@dataclass
class Suministrador:
    """CAP-014 suministrador row attached to an Expediente."""

    id: UUID
    id_expediente: UUID
    id_suministrador: UUID
    id_padre: UUID | None
    descripcion: str | None
    contratista_principal: bool | None
    sub_contratista: bool | None
    created_at: datetime | None = None
    updated_at: datetime | None = None

    def __post_init__(self) -> None:
        if self.id is None or self.id_expediente is None or self.id_suministrador is None:
            raise ValueError("id, id_expediente and id_suministrador are required")
        if self.id_padre == self.id:
            raise ValueError("id_padre cannot be the row itself (cycle detected)")
        if self.descripcion is not None and len(self.descripcion) > MAX_DESCRIPCION_LEN:
            raise ValueError(f"descripcion must not exceed {MAX_DESCRIPCION_LEN} characters")
