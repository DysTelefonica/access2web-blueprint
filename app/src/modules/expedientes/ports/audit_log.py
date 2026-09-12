"""AuditLogPort — D-EXP-3, D-EXP-4, D55."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Protocol
from uuid import UUID


@dataclass(frozen=True)
class ExpedienteAuditEvent:
    """Audit event. Stripped of telemetry per D55."""

    id: UUID
    event_type: str
    actor_id: UUID
    target_id: UUID
    capacidad: str  # e.g. "EXP-CAP-001"
    module: str = "expedientes"
    result: str = "ok"  # "ok" | "denied" | "error"
    correlation_id: UUID | None = None
    payload: dict[str, Any] = field(default_factory=dict)
    created_at: datetime | None = None


class AuditLogPort(Protocol):
    """Append-only audit-log boundary."""

    async def append(self, event: ExpedienteAuditEvent) -> None: ...
    async def list_for_actor(
        self, actor_id: UUID, since: datetime
    ) -> Sequence[ExpedienteAuditEvent]: ...
