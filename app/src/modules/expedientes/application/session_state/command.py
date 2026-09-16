"""Commands and outcomes for EXP-CAP-045 Estado de sesión adapter."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID

EXPEDIENTES_APP_ID: int = 19


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
    app_id: int = EXPEDIENTES_APP_ID


__all__ = [
    "EXPEDIENTES_APP_ID",
    "SessionStateAuthorizationError",
    "SessionStateCommand",
    "SessionStateError",
    "SessionStateValidationError",
    "SessionState",
]
