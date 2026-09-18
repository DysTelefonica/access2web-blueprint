"""Audit evidence builder for EXP-CAP-036 E2E hash."""

from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.domain.hash.versioning import HashVersioned
from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent


def hash_event(actor_id: UUID, result: HashVersioned) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="e2e.hash.computed",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-036",
        created_at=datetime.now(UTC),
        payload={"algorithm": result.algorithm, "digest": result.digest},
    )


__all__ = ["hash_event"]
