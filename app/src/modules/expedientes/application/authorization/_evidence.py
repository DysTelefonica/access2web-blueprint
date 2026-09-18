"""Audit evidence builder for EXP-CAP-043 authorization policy."""

from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.authorization.command import (
    AuthorizationResult,
)
from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent


def authz_event(actor_id: UUID, result: AuthorizationResult) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type=("e2e.authz.allowed" if result.allowed else "e2e.authz.denied"),
        actor_id=actor_id,
        target_id=actor_id,
        capacidad=result.capability,
        created_at=datetime.now(UTC),
        payload={"allowed": result.allowed, "denied_reason": result.denied_reason},
    )


__all__ = ["authz_event"]
