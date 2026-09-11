"""ExpedienteEstado StrEnum — D-EXP-1, CAP-003..CAP-005."""

from __future__ import annotations

from enum import StrEnum


class ExpedienteEstado(StrEnum):
    """Expediente lifecycle states.

    Values match the Postgres ENUM literal so JSON serialisation round-trips
    without a custom encoder.
    """

    BORRADOR = "BORRADOR"
    ADJUDICADO = "ADJUDICADO"
    FORMALIZADO = "FORMALIZADO"
    ARCHIVADO = "ARCHIVADO"
