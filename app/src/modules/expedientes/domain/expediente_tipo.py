"""ExpedienteTipo StrEnum — D-EXP-1, CAP-006.

AM = Acuerdo Marco. LOTE = Lote dentro de un AM.
BASED = Expediente basado en un AM/Lote.
"""

from __future__ import annotations

from enum import StrEnum


class ExpedienteTipo(StrEnum):
    """Expediente type classification.

    AM: Acuerdo Marco (acuerdo padre, no tiene padre propio).
    LOTE: Lote dentro de un AM (requiere id_expediente_padre).
    BASED: Expediente basado en un AM o Lote (requiere id_expediente_padre).
    """

    AM = "AM"
    LOTE = "LOTE"
    BASED = "BASED"
