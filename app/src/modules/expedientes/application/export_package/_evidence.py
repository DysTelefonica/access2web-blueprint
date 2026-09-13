"""Audit evidence builder for EXP-CAP-037 E2E package."""

from datetime import UTC, datetime
from typing import Any
from uuid import UUID, uuid4

from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent


def package_event(actor_id: UUID, package: dict[str, Any]) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="e2e.package.assembled",
        actor_id=actor_id,
        target_id=actor_id,
        capacidad="EXP-CAP-037",
        created_at=datetime.now(UTC),
        payload={
            "package_id": str(package["package_id"]),
            "api_version": package.get("api_version"),
            "item_count": len(package.get("items", [])),
        },
    )


__all__ = ["package_event"]
