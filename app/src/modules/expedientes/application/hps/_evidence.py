"""Audit evidence builder for EXP-CAP-051 HPS adapter (H01).

Canonical home for ``EXPEDIENTES_APP_ID = 19``; keeps ``command.py``
free of cross-file duplication (check_dry BASELINE ``dup:79d8c5976f0f``).
"""

from collections.abc import Mapping
from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent


def hps_event(
    action: str,
    *,
    actor_id: UUID,
    target_id: UUID,
    payload: Mapping[str, object],
    correlation_id: UUID | None = None,
) -> ExpedienteAuditEvent:
    """Build the ``ExpedienteAuditEvent`` for a HPS submit or query action."""
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type=f"hps.{action}.recorded",
        actor_id=actor_id,
        target_id=target_id,
        capacidad="EXP-CAP-051",
        module="expedientes",
        result="ok",
        correlation_id=correlation_id,
        payload=dict(payload) | {"app_id": EXPEDIENTES_APP_ID},
        created_at=datetime.now(UTC),
    )


EXPEDIENTES_APP_ID: int = 19


__all__ = ["EXPEDIENTES_APP_ID", "hps_event"]
