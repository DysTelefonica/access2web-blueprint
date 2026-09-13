"""Audit evidence builder for EXP-CAP-034 canonical JSON."""

from datetime import datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.json_canonical.command import (
    CANONICAL_API_VERSION,
)
from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent


def render_event(actor_id: UUID, at: datetime, total_rows: int) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="e2e.canonical.rendered",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-034",
        created_at=at,
        payload={"api_version": CANONICAL_API_VERSION, "total_rows": total_rows},
    )


__all__ = ["render_event"]
