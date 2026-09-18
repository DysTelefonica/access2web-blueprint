"""Responsable child entity — D-EXP-1, CAP-012.

Legacy ``TbExpedientesResponsables`` carries:

- IDExpedienteResponsable (PK)
- IdExpediente (FK)
- IdUsuario (FK)
- CorreoSiempre / EsJefeProyecto / esPreventa (Text(2), nullable, Yes/No)

The D102 cross-cutting rule will standardise these to BOOLEAN in
PostgreSQL, but the domain keeps them nullable today so legacy rows
remain valid. ``rol`` is the user-supplied application tag; the spec
only constrains its length.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID

MAX_ROL_LEN = 64


@dataclass
class Responsable:
    """CAP-012 responsable attached to an Expediente."""

    id: UUID
    id_expediente: UUID
    id_usuario: UUID
    rol: str | None
    correo_siempre: bool | None
    es_jefe_proyecto: bool | None
    es_preventa: bool | None
    created_at: datetime | None = None
    updated_at: datetime | None = None

    def __post_init__(self) -> None:
        if self.id is None or self.id_expediente is None or self.id_usuario is None:
            raise ValueError("id, id_expediente and id_usuario are required")
        if self.rol is not None:
            cleaned = self.rol.strip()
            if not cleaned:
                raise ValueError("rol must not be blank")
            if len(cleaned) > MAX_ROL_LEN:
                raise ValueError(f"rol must not exceed {MAX_ROL_LEN} characters")
            self.rol = cleaned
