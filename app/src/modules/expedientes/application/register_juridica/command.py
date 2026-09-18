"""Commands and outcomes for EXP-CAP-013 juridica registration."""

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID


class RegisterJuridicaError(Exception):
    """Base error for juridica registration."""


class RegisterJuridicaValidationError(RegisterJuridicaError):
    """The command or referenced expediente is invalid."""


class RegisterJuridicaAuthorizationError(RegisterJuridicaError):
    """The caller has no authenticated actor."""


class RegisterJuridicaConflictError(RegisterJuridicaError):
    """The juridica identity was already used for different data."""


@dataclass(frozen=True)
class RegisterJuridicaCommand:
    juridica_id: UUID
    expediente_id: UUID
    id_juridica: UUID
    id_suministrador: UUID | None
    contratista_principal: bool | None
    sub_contratista: bool | None
    actor_id: UUID


@dataclass(frozen=True)
class RegisterJuridicaResult:
    juridica_id: UUID
    expediente_id: UUID
    registered_at: datetime
