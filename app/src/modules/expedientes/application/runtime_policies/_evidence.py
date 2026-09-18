"""Audit evidence builder for EXP-CAP-047..050 (T01).

Canonical home for ``EXPEDIENTES_APP_ID = 19``; keeps ``command.py``
free of cross-file duplication (check_dry BASELINE ``dup:79d8c5976f0f``).
"""

from collections.abc import Mapping
from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent

EXPEDIENTES_APP_ID: int = 19

_CAP_MAP: dict[str, str] = {
    "readiness": "EXP-CAP-047",
    "config": "EXP-CAP-048",
    "cache": "EXP-CAP-049",
    "binding": "EXP-CAP-050",
}


def runtime_event(
    category: str,
    *,
    actor_id: UUID,
    payload: Mapping[str, object],
) -> ExpedienteAuditEvent:
    """Build the ``ExpedienteAuditEvent`` for one of the four runtime capabilities."""
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type=f"runtime.{category}.recorded",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad=_CAP_MAP[category],
        module="expedientes",
        result="ok",
        correlation_id=None,
        payload={**payload, "app_id": EXPEDIENTES_APP_ID},
        created_at=datetime.now(UTC),
    )


__all__ = ["EXPEDIENTES_APP_ID", "runtime_event"]
