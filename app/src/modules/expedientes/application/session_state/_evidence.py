"""Audit evidence builder for EXP-CAP-045 Estado de sesión adapter."""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.session_state.command import (
    EXPEDIENTES_APP_ID,
)
from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent


def session_bound_event(actor_id: UUID, session_id: UUID) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="authz.session.bound",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-045",
        result="ok",
        payload={"session_id": session_id, "app_id": EXPEDIENTES_APP_ID},
        created_at=datetime.now(UTC),
    )


def session_failed_event(
    actor_id: UUID, reason: str, *, error_class: str | None = None
) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="authz.session.failed",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-045",
        result="error",
        payload={"reason": reason, "error_class": error_class},
        created_at=datetime.now(UTC),
    )


def session_cleared_event(actor_id: UUID, session_id: UUID) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="authz.session.cleared",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-045",
        result="ok",
        payload={"session_id": session_id, "app_id": EXPEDIENTES_APP_ID},
        created_at=datetime.now(UTC),
    )


__all__ = ["session_bound_event", "session_cleared_event", "session_failed_event"]
