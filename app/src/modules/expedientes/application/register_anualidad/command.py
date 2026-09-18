"""Commands and outcomes for EXP-CAP-011 anualidades."""

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from uuid import UUID


class RegisterAnualidadError(Exception):
    """Base error for anualidad registration."""


class RegisterAnualidadValidationError(RegisterAnualidadError):
    """The command or referenced expediente is invalid."""


class RegisterAnualidadAuthorizationError(RegisterAnualidadError):
    """The caller has no authenticated actor."""


class RegisterAnualidadConflictError(RegisterAnualidadError):
    """The anualidad identity was already used for different data."""


@dataclass(frozen=True)
class RegisterAnualidadCommand:
    anualidad_id: UUID
    expediente_id: UUID
    anio: int | None
    bi_iva: Decimal | None
    bi_ipsi: Decimal | None
    bi_igic: Decimal | None
    bi_exenta: Decimal | None
    iva: Decimal | None
    ipsi: Decimal | None
    igic: Decimal | None
    periodo_facturacion: str | None
    actor_id: UUID


@dataclass(frozen=True)
class RegisterAnualidadResult:
    anualidad_id: UUID
    expediente_id: UUID
    registered_at: datetime
