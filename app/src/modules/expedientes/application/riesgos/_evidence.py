"""Audit evidence builder for EXP-CAP-053 Riesgos adapter (H03).

Canonical home for ``EXPEDIENTES_APP_ID = 19`` (DRY ratchet
``dup:79d8c5976f0f``).
"""

from collections.abc import Mapping
from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent

EXPEDIENTES_APP_ID: int = 19


def riesgos_event(
    *,
    actor_id: UUID,
    target_id: UUID,
    payload: Mapping[str, object],
) -> ExpedienteAuditEvent:
    """Build the ``ExpedienteAuditEvent`` for a Riesgos lookup."""
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="riesgos.lookup.recorded",
        actor_id=actor_id,
        target_id=target_id,
        capacidad="EXP-CAP-053",
        module="expedientes",
        result="ok",
        correlation_id=None,
        payload=dict(payload) | {"app_id": EXPEDIENTES_APP_ID},
        created_at=datetime.now(UTC),
    )


__all__ = ["EXPEDIENTES_APP_ID", "riesgos_event"]
