"""Command and outcomes for EXP-CAP-009 modification history."""

from dataclasses import dataclass
from datetime import date, datetime
from uuid import UUID


class RegisterModificadoError(Exception):
    """Base error for modification history registration."""


class RegisterModificadoValidationError(RegisterModificadoError):
    """The command or referenced expediente is invalid."""


class RegisterModificadoAuthorizationError(RegisterModificadoError):
    """The caller has no authenticated actor."""


class RegisterModificadoConflictError(RegisterModificadoError):
    """The modification entry identity was already used for different data."""


@dataclass(frozen=True)
class RegisterModificadoCommand:
    modificado_id: UUID
    expediente_id: UUID
    n_modificado: str | None
    fecha_firma: date | None
    fecha_fin: date | None
    descripcion: str | None
    actor_id: UUID


@dataclass(frozen=True)
class RegisterModificadoResult:
    modificado_id: UUID
    expediente_id: UUID
    registered_at: datetime | None
