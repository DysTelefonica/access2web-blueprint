"""Audit evidence builder for EXP-CAP-042 E2E ordinal expansion."""

from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.ordinal_e2e.command import (
    OrdinalExpansionResult,
)
from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent


def ordinal_event(actor_id: UUID, result: OrdinalExpansionResult) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="e2e.ordinal.assigned",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-042",
        created_at=datetime.now(UTC),
        payload={
            "root_id": str(result.root_ordinal) if result.root_ordinal else None,
            "assignments": [
                {"node_id": str(a.node_id), "ordinal": a.ordinal} for a in result.assignments
            ],
        },
    )


__all__ = ["ordinal_event"]
