"""Audit evidence builder for EXP-CAP-035 E2E batch."""

from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.batch_e2e.command import (
    BatchE2EResult,
)
from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent


def batch_event(actor_id: UUID, result: BatchE2EResult) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="e2e.batch.executed",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-035",
        created_at=datetime.now(UTC),
        payload={
            "batch_id": str(result.batch_id),
            "status": result.status.value,
            "detail_count": result.detail_count,
            "succeeded_count": result.succeeded_count,
            "failed_count": result.failed_count,
            "atomicity_rolled_back": result.atomicity_rolled_back,
        },
    )


__all__ = ["batch_event"]
