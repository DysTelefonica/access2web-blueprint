"""Commands and outcomes for EXP-CAP-041 E2E session."""

from dataclasses import dataclass
from datetime import datetime
from enum import StrEnum
from uuid import UUID


class E2ESessionError(Exception):
    """Base error for the E2E session use case."""


class E2ESessionValidationError(E2ESessionError):
    """The session id is unknown or the state does not allow the action."""


class E2ESessionAuthorizationError(E2ESessionError):
    """The caller has no authenticated actor, no permission or no ownership."""


class E2ESessionStatus(StrEnum):
    """Lifecycle of an E2E session."""

    ACTIVE = "active"
    CLOSED = "closed"


@dataclass(frozen=True)
class E2ESessionCommand:
    actor_id: UUID | None
    session_id: UUID | None = None


@dataclass(frozen=True)
class E2ESessionResult:
    """The result of an E2E session action."""

    session_id: UUID
    actor_id: UUID
    status: E2ESessionStatus
    opened_at: datetime
    last_seen_at: datetime


__all__ = [
    "E2ESessionAuthorizationError",
    "E2ESessionCommand",
    "E2ESessionError",
    "E2ESessionResult",
    "E2ESessionStatus",
    "E2ESessionValidationError",
]
