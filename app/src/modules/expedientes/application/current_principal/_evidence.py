"""Audit evidence builder for EXP-CAP-044 CurrentPrincipal adapter."""

from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.current_principal.command import (
    EXPEDIENTES_APP_ID,
    Principal,
)
from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent


def principal_loaded_event(actor_id: UUID, principal: Principal) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="authz.principal.loaded",
        actor_id=actor_id,
        target_id=principal.user_id,
        capacidad="EXP-CAP-044",
        result="ok",
        payload={
            "app_id": principal.app_id,
            "permissions_count": len(principal.permissions),
            "permissions": sorted(principal.permissions),
        },
        created_at=datetime.now(UTC),
    )


def principal_failed_event(
    actor_id: UUID, reason: str, *, error_class: str | None = None
) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="authz.principal.failed",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-044",
        result="error",
        payload={
            "reason": reason,
            "error_class": error_class,
            "app_id": EXPEDIENTES_APP_ID,
        },
        created_at=datetime.now(UTC),
    )


__all__ = ["principal_loaded_event", "principal_failed_event"]
