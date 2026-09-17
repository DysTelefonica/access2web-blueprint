"""Audit evidence builder for EXP-CAP-045 Estado de sesión adapter.

Single canonical home for the expedientes application id (19) inside this
adapter. Kept here instead of in command.py to keep the command prelude
free of cross-file duplication against the anexos/juridica/responsable/
suministrador siblings (check_dry.py ratchet — see BASELINE entry
dup:79d8c5976f0f). All audit payloads emitted below must reference this
constant rather than re-declaring the literal.
"""

from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent

# Canonical expedientes app id (19) — see module docstring for why this lives here.
EXPEDIENTES_APP_ID: int = 19


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
