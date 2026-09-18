"""Commands and outcomes for EXP-CAP-043 authorization policy."""

from dataclasses import dataclass
from uuid import UUID

EFFECTIVE_DENY: str = "deny_by_default"


class AuthorizationError(Exception):
    """Base error for the authorization use case."""


class AuthorizationValidationError(AuthorizationError):
    """The capability is missing or empty."""


class AuthorizationAuthorizationError(AuthorizationError):
    """The caller has no authenticated actor or the actor is empty."""


@dataclass(frozen=True)
class AuthorizationCommand:
    actor_id: UUID | None
    capability: str
    permissions: set[str]


@dataclass(frozen=True)
class AuthorizationResult:
    """The outcome of an authorization decision (CAP-043)."""

    actor_id: UUID
    capability: str
    allowed: bool
    denied_reason: str | None


__all__ = [
    "AuthorizationAuthorizationError",
    "AuthorizationCommand",
    "AuthorizationError",
    "AuthorizationResult",
    "AuthorizationValidationError",
    "EFFECTIVE_DENY",
]
