"""Audit evidence builder for EXP-CAP-046 (A04).

Canonical home for ``EXPEDIENTES_APP_ID = 19``; keeps ``command.py`` free of
cross-file duplication (check_dry BASELINE ``dup:79d8c5976f0f``).
"""

from datetime import UTC, datetime
from typing import Any
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.audit.command import AuditCategory
from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent

# Canonical expedientes app id (19) — see module docstring for why this lives here.
EXPEDIENTES_APP_ID: int = 19


def audit_event_for(
    category: AuditCategory,
    *,
    actor_id: UUID,
    target_id: UUID,
    correlation_id: UUID | None,
    payload: dict[str, Any],
) -> ExpedienteAuditEvent:
    """Build the ``ExpedienteAuditEvent`` for one of the four audit categories.

    Caller (``AuditService``) validates ``actor_id`` upstream; this builder
    assumes a non-empty UUID.
    """
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type=f"audit.{category}.recorded",
        actor_id=actor_id,
        target_id=target_id,
        capacidad="EXP-CAP-046",
        module="expedientes",
        result="ok",
        correlation_id=correlation_id,
        payload={**payload, "app_id": EXPEDIENTES_APP_ID},
        created_at=datetime.now(UTC),
    )


__all__ = ["EXPEDIENTES_APP_ID", "audit_event_for"]
