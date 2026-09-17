"""Commands, outcomes and error hierarchy for EXP-CAP-046 (A04).

``EXPEDIENTES_APP_ID`` lives in ``_evidence.py`` (not here) — see the module
docstring there for the DRY-ratchet rationale (``dup:79d8c5976f0f``).
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any, Literal
from uuid import UUID

AuditCategory = Literal["access", "change", "export", "integration"]


class AuditError(Exception):
    """Base error for the audit use case."""


class AuditAuthorizationError(AuditError):
    """Deny-by-default: raised before any audit emission."""


class AuditDependencyError(AuditError):
    """Wraps the underlying driver exception (``__cause__``) so the caller's
    UoW can roll back without losing the dependency-failure context.
    """


@dataclass(frozen=True)
class AuditCommand:
    """Input for ``AuditService.record_*``. ``actor_id`` is typed ``UUID | None``
    so the deny-by-default guard is statically checkable.
    """

    actor_id: UUID | None
    target_id: UUID
    correlation_id: UUID | None
    payload: dict[str, Any]


@dataclass(frozen=True)
class AuditResult:
    """Outcome of a successful audit emission (CAP-046)."""

    actor_id: UUID
    category: AuditCategory
    event_id: UUID
    recorded_at: datetime


__all__ = [
    "AuditAuthorizationError",
    "AuditCategory",
    "AuditCommand",
    "AuditDependencyError",
    "AuditError",
    "AuditResult",
]
