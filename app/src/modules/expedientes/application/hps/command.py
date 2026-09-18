"""Commands, results and error hierarchy for EXP-CAP-051 HPS adapter (H01).

``EXPEDIENTES_APP_ID`` lives in ``_evidence.py`` (not here) — see the
DRY-ratchet rationale (``dup:79d8c5976f0f``).
"""

from dataclasses import dataclass
from typing import Any
from uuid import UUID


class HpsError(Exception):
    """Base error for the HPS use case."""


class HpsAuthorizationError(HpsError):
    """Deny-by-default: raised before any audit emission or adapter call."""


class HpsDependencyError(HpsError):
    """Wraps the underlying driver exception (``__cause__``) so the
    caller's UoW can roll back without losing the dependency-failure
    context.
    """


@dataclass(frozen=True)
class HpsSubmitCommand:
    actor_id: UUID | None
    idempotency_key: UUID
    expediente_id: UUID
    payload: dict[str, Any]
    credential: str


@dataclass(frozen=True)
class HpsQueryCommand:
    actor_id: UUID | None
    reference: UUID
    credential: str


@dataclass(frozen=True)
class HpsResult:
    actor_id: UUID
    reference: UUID
    status: str
    deduped: bool
    payload: dict[str, Any]


__all__ = [
    "HpsAuthorizationError",
    "HpsDependencyError",
    "HpsError",
    "HpsQueryCommand",
    "HpsResult",
    "HpsSubmitCommand",
]
