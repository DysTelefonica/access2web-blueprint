"""Commands and outcomes for EXP-CAP-044 CurrentPrincipal adapter."""

from dataclasses import dataclass
from uuid import UUID

EXPEDIENTES_APP_ID: int = 19


class CurrentPrincipalError(Exception):
    """Base error for the CurrentPrincipal use case."""


class CurrentPrincipalValidationError(CurrentPrincipalError):
    """The command is missing or invalid."""


class CurrentPrincipalAuthorizationError(CurrentPrincipalError):
    """The caller has no authenticated actor or the actor is empty."""


@dataclass(frozen=True)
class CurrentPrincipalCommand:
    actor_id: UUID | None


@dataclass(frozen=True)
class Principal:
    """The resolved identity and permissions for the current actor (CAP-044)."""

    user_id: UUID
    app_id: int
    permissions: frozenset[str]

    @classmethod
    def deny(cls, user_id: UUID, *, app_id: int = EXPEDIENTES_APP_ID) -> "Principal":
        """Build a deny-by-default Principal with an empty permission set."""
        return cls(user_id=user_id, app_id=app_id, permissions=frozenset())


@dataclass(frozen=True)
class CurrentPrincipalResult:
    """The outcome of a CurrentPrincipal resolution (CAP-044)."""

    principal: Principal
    loaded: bool


__all__ = [
    "CurrentPrincipalAuthorizationError",
    "CurrentPrincipalCommand",
    "CurrentPrincipalError",
    "CurrentPrincipalResult",
    "CurrentPrincipalValidationError",
    "EXPEDIENTES_APP_ID",
    "Principal",
]
