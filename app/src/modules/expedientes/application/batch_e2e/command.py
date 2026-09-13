"""Commands and outcomes for EXP-CAP-035 E2E batch atomicity."""

from dataclasses import dataclass, field
from enum import StrEnum
from typing import Any
from uuid import UUID


class BatchE2EError(Exception):
    """Base error for the E2E batch use case."""


class BatchE2EValidationError(BatchE2EError):
    """The batch is empty, the atomicity is unknown, or the input is malformed."""


class BatchE2EAuthorizationError(BatchE2EError):
    """The caller has no authenticated actor or lacks permission."""


class BatchStatus(StrEnum):
    """Lifecycle of an E2E batch."""

    PENDING = "pending"
    COMPLETED = "completed"
    FAILED = "failed"


class BatchAtomicity(StrEnum):
    """Atomicity policy applied when one or more details fail."""

    ALL_OR_NOTHING = "all_or_nothing"
    BEST_EFFORT = "best_effort"


@dataclass(frozen=True)
class BatchDetailInput:
    id: UUID
    payload: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class BatchE2ECommand:
    actor_id: UUID | None
    details: list[BatchDetailInput]
    atomicity: str = "all_or_nothing"


@dataclass(frozen=True)
class BatchE2EResult:
    """The outcome of an E2E batch execution."""

    batch_id: UUID
    status: BatchStatus
    detail_count: int
    succeeded_count: int
    failed_count: int
    atomicity_rolled_back: bool


__all__ = [
    "BatchAtomicity",
    "BatchDetailInput",
    "BatchE2EAuthorizationError",
    "BatchE2ECommand",
    "BatchE2EError",
    "BatchE2EResult",
    "BatchE2EValidationError",
    "BatchStatus",
]
