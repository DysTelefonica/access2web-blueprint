"""Service implementation for EXP-CAP-046 (A04).

``AuditService.record_change`` emits the audit event for a mutation — it does
NOT call ``audit_log.record_change``. The per-field ``ChangeRecord`` row in
``TbCambios`` is the caller's UoW responsibility (CAP-001 §Camino feliz, DA-11).

Duck-typed ``audit_log: Any`` — a Protocol would force a cross-module import
that the layer gate (``check_layers.py``) flags as application → driver.
"""

from __future__ import annotations

from typing import Any
from uuid import UUID

from app.src.modules.expedientes.application.audit._evidence import (
    audit_event_for,
)
from app.src.modules.expedientes.application.audit.command import (
    AuditAuthorizationError,
    AuditCategory,
    AuditCommand,
    AuditDependencyError,
    AuditResult,
)


class AuditService:
    """Application-level audit-event emitter (CAP-046). One event per call."""

    def __init__(self, *, audit_log: Any) -> None:
        self._audit_log = audit_log

    async def record_access(self, command: AuditCommand) -> AuditResult:
        return await self._record("access", command)

    async def record_change(self, command: AuditCommand) -> AuditResult:
        return await self._record("change", command)

    async def record_export(self, command: AuditCommand) -> AuditResult:
        return await self._record("export", command)

    async def record_integration(self, command: AuditCommand) -> AuditResult:
        return await self._record("integration", command)

    async def _record(self, category: AuditCategory, command: AuditCommand) -> AuditResult:
        actor_id = self._validated_actor(command)
        event = audit_event_for(
            category,
            actor_id=actor_id,
            target_id=command.target_id,
            correlation_id=command.correlation_id,
            payload=command.payload,
        )
        try:
            await self._audit_log.append(event)
        except Exception as exc:  # noqa: BLE001 — wrap any driver failure
            raise AuditDependencyError(
                "audit_log.append failed; caller UoW must roll back"
            ) from exc
        recorded_at = event.created_at
        assert recorded_at is not None, "audit_event_for must stamp created_at"
        return AuditResult(
            actor_id=actor_id,
            category=category,
            event_id=event.id,
            recorded_at=recorded_at,
        )

    @staticmethod
    def _validated_actor(command: AuditCommand) -> UUID:
        """Deny-by-default: ``None`` and ``UUID(int=0)`` are rejected before any
        audit emission so the log is never polluted by the rejected request.
        """
        if command.actor_id is None or command.actor_id == UUID(int=0):
            raise AuditAuthorizationError("actor_id is required (deny-by-default)")
        return command.actor_id


__all__ = ["AuditService"]
