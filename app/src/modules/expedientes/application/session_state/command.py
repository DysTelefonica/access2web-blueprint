"""Commands and outcomes for EXP-CAP-045 Estado de sesión adapter (A03)."""

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID


class SessionStateError(Exception):
    """Base error for the SessionState use case."""


class SessionStateAuthorizationError(SessionStateError):
    """The caller has no authenticated actor or the actor is empty."""


class SessionStateValidationError(SessionStateError):
    """The command is missing or invalid."""


@dataclass(frozen=True)
class SessionStateCommand:
    actor_id: UUID | None


@dataclass(frozen=True)
class SessionState:
    """Per-request session bound to an authenticated actor (CAP-045)."""

    actor_id: UUID
    session_id: UUID
    bound_at: datetime
    app_id: int = 19  # canonical expedientes app id (see _evidence.py)


__all__ = [
    "SessionStateAuthorizationError",
    "SessionStateCommand",
    "SessionStateError",
    "SessionStateValidationError",
    "SessionState",
]
