"""Anualidad child entity — D-EXP-1, CAP-011.

Legacy ``TbExpedientesAnualidades`` carries:

- IDAnualidad (PK)
- IDExpediente (FK, nullable)
- Año (small integer)
- BIIVA / BIIPSI / BIIGIC / BIEXENTA — tax bases
- IVA / IPSI / IGIC — tax quotes
- PeriodoFacturacion (255, nullable)

D-EXP-2 keeps every column nullable in the legacy schema. The
domain preserves nulls and only enforces the invariants the
regulatory spec demands: the year range, non-negative amounts and
the period label length.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from uuid import UUID

MIN_ANIO = 1970
MAX_ANIO = 2100


@dataclass
class Anualidad:
    """CAP-011 anualidad attached to an Expediente."""

    id: UUID
    id_expediente: UUID
    anio: int | None
    bi_iva: Decimal | None
    bi_ipsi: Decimal | None
    bi_igic: Decimal | None
    bi_exenta: Decimal | None
    iva: Decimal | None
    ipsi: Decimal | None
    igic: Decimal | None
    periodo_facturacion: str | None
    created_at: datetime | None = None
    updated_at: datetime | None = None

    def __post_init__(self) -> None:
        if self.id is None or self.id_expediente is None:
            raise ValueError("id and id_expediente are required")
        self._validate_anio()
        self._validate_non_negative("bi_iva", self.bi_iva)
        self._validate_non_negative("bi_ipsi", self.bi_ipsi)
        self._validate_non_negative("bi_igic", self.bi_igic)
        self._validate_non_negative("bi_exenta", self.bi_exenta)
        self._validate_non_negative("iva", self.iva)
        self._validate_non_negative("ipsi", self.ipsi)
        self._validate_non_negative("igic", self.igic)
        self._validate_periodo()

    def _validate_anio(self) -> None:
        if self.anio is None:
            return
        if not isinstance(self.anio, int) or isinstance(self.anio, bool):
            raise ValueError("anio must be an integer")
        if self.anio < MIN_ANIO or self.anio > MAX_ANIO:
            raise ValueError(f"anio must be between {MIN_ANIO} and {MAX_ANIO}, got {self.anio}")

    def _validate_periodo(self) -> None:
        if self.periodo_facturacion is None:
            return
        if not isinstance(self.periodo_facturacion, str):
            raise ValueError("periodo_facturacion must be a string")
        if len(self.periodo_facturacion) > 255:
            raise ValueError("periodo_facturacion must not exceed 255 characters")

    @staticmethod
    def _validate_non_negative(field: str, value: Decimal | None) -> None:
        if value is None:
            return
        if not isinstance(value, Decimal):
            raise ValueError(f"{field} must be a Decimal")
        if value < 0:
            raise ValueError(f"{field} must be non-negative")
