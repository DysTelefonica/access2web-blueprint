"""Commands and outcomes for EXP-CAP-032 feedback guard."""

from dataclasses import dataclass
from datetime import datetime
from enum import StrEnum
from typing import Any
from uuid import UUID


class FeedbackGuardError(Exception):
    """Base error for the feedback guard."""


class FeedbackGuardValidationError(FeedbackGuardError):
    """The command is missing required fields."""


class FeedbackGuardAuthorizationError(FeedbackGuardError):
    """The caller has no authenticated actor."""


class FeedbackGuardStatus(StrEnum):
    """Lifecycle states exposed to the UI as ``aria-busy`` / ``aria-live``.

    - ``BUSY``: the operation was accepted and is in flight; the
      delivery layer renders the button as busy and suppresses the
      double-submit guard.
    - ``FINISHED``: a replay returned the recorded result; the
      delivery layer can clear the busy state and render success.
    - ``ERROR``: the operation failed terminally; the delivery layer
      must surface the error and re-enable the form.
    """

    BUSY = "busy"
    FINISHED = "finished"
    ERROR = "error"


@dataclass(frozen=True)
class FeedbackGuardCommand:
    idempotency_key: UUID
    operation: str
    actor_id: UUID


@dataclass(frozen=True)
class FeedbackGuardState:
    """Result returned to the delivery layer.

    ``completed_at`` is set on a FINISHED replay (it is the wall-clock
    instant when the original operation committed). ``result`` carries
    whatever payload the originating use case recorded — the feedback
    guard is opaque to that data.
    """

    idempotency_key: UUID
    operation: str
    status: FeedbackGuardStatus
    completed_at: datetime | None
    result: Any | None

    def is_terminal(self) -> bool:
        return self.status is not FeedbackGuardStatus.BUSY
