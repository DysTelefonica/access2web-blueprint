"""Commands and outcomes for EXP-CAP-012 responsable assignment."""

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID


class RegisterResponsableError(Exception):
    """Base error for responsable assignment."""


class RegisterResponsableValidationError(RegisterResponsableError):
    """The command or referenced expediente is invalid."""


class RegisterResponsableAuthorizationError(RegisterResponsableError):
    """The caller has no authenticated actor."""


class RegisterResponsableConflictError(RegisterResponsableError):
    """The responsable identity was already used for different data."""


@dataclass(frozen=True)
class RegisterResponsableCommand:
    responsable_id: UUID
    expediente_id: UUID
    usuario_id: UUID
    rol: str | None
    correo_siempre: bool | None
    es_jefe_proyecto: bool | None
    es_preventa: bool | None
    actor_id: UUID


@dataclass(frozen=True)
class RegisterResponsableResult:
    responsable_id: UUID
    expediente_id: UUID
    registered_at: datetime
