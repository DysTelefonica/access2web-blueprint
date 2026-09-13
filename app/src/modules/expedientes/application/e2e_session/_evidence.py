"""Audit evidence builder for EXP-CAP-041 E2E session."""

from datetime import UTC, datetime
from uuid import UUID

from app.src.modules.expedientes.ports.audit_log import (
    ChangeRecord,
    ExpedienteAuditEvent,
)


def session_event(
    actor_id: UUID,
    event_type: str,
    session_id: UUID,
) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=__import__("uuid").uuid4(),
        event_type=event_type,
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-041",
        created_at=datetime.now(UTC),
        payload={"session_id": str(session_id)},
    )


def session_change(
    actor_id: UUID,
    session_id: UUID,
    status: str,
) -> ChangeRecord:
    return ChangeRecord(
        id=__import__("uuid").uuid4(),
        nombre_tabla="e2e_sessions",
        id_expediente=__import__("uuid").UUID(int=0),
        nombre_campo="status",
        valor_inicial=None,
        valor_final=status,
        fecha_cambio=datetime.now(UTC),
        id_usuario_cambio=actor_id,
        accion="edit",
    )


__all__ = ["session_change", "session_event"]
