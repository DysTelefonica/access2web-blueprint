"""Command and outcomes for EXP-CAP-008 milestone registration."""

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from uuid import UUID


class RegisterHitoError(Exception):
    """Base error for milestone registration."""


class RegisterHitoValidationError(RegisterHitoError):
    """The command or referenced expediente is invalid."""


class RegisterHitoAuthorizationError(RegisterHitoError):
    """The caller has no authenticated actor."""


class RegisterHitoConflictError(RegisterHitoError):
    """The milestone identity was already used for different data."""


@dataclass(frozen=True)
class RegisterHitoCommand:
    hito_id: UUID
    expediente_id: UUID
    descripcion: str | None
    fecha_hito: date
    garantia_fecha_fin: date | None
    importe: Decimal | None
    actor_id: UUID


@dataclass(frozen=True)
class RegisterHitoResult:
    hito_id: UUID
    expediente_id: UUID
    registered_at: datetime | None
