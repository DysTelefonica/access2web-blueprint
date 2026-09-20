"""Audit evidence builder for EXP-CAP-054 NC adapter (H04)."""

from collections.abc import Mapping
from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent

EXPEDIENTES_APP_ID: int = 19


def nc_event(
    *,
    actor_id: UUID,
    target_id: UUID,
    payload: Mapping[str, object],
) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="nc.lookup.recorded",
        actor_id=actor_id,
        target_id=target_id,
        capacidad="EXP-CAP-054",
        module="expedientes",
        result="ok",
        correlation_id=None,
        payload=dict(payload) | {"app_id": EXPEDIENTES_APP_ID},
        created_at=datetime.now(UTC),
    )


__all__ = ["EXPEDIENTES_APP_ID", "nc_event"]
