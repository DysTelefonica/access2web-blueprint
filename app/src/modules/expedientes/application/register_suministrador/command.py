"""Commands and outcomes for EXP-CAP-014 suministrador registration."""

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID


class RegisterSuministradorError(Exception):
    """Base error for suministrador registration."""


class RegisterSuministradorValidationError(RegisterSuministradorError):
    """The command or referenced expediente is invalid."""


class RegisterSuministradorAuthorizationError(RegisterSuministradorError):
    """The caller has no authenticated actor."""


class RegisterSuministradorConflictError(RegisterSuministradorError):
    """The suministrador identity was already used for different data."""


@dataclass(frozen=True)
class RegisterSuministradorCommand:
    suministrador_id: UUID
    expediente_id: UUID
    id_suministrador: UUID
    id_padre: UUID | None
    descripcion: str | None
    contratista_principal: bool | None
    sub_contratista: bool | None
    actor_id: UUID


@dataclass(frozen=True)
class RegisterSuministradorResult:
    suministrador_id: UUID
    expediente_id: UUID
    registered_at: datetime
