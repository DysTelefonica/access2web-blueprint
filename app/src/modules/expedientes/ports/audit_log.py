"""AuditLogPort — D-EXP-3, D-EXP-4, D55."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Protocol
from uuid import UUID


@dataclass(frozen=True)
class ExpedienteAuditEvent:
    """Audit event. Stripped of telemetry per D55."""

    id: UUID
    event_type: str
    actor_id: UUID
    target_id: UUID
    capacidad: str  # e.g. "EXP-CAP-001"
    module: str = "expedientes"
    result: str = "ok"  # "ok" | "denied" | "error"
    correlation_id: UUID | None = None
    payload: dict[str, Any] = field(default_factory=dict)
    created_at: datetime | None = None


@dataclass(frozen=True)
class ChangeRecord:
    """Row for ``TbCambios`` — the historical log of mutations.

    CAP-001 §Camino feliz requires the last-change row to be written
    inside the same UoW as the cabecera. This value object carries
    the fields needed for one row in ``TbCambios`` (one row per
    mutation). For a high field-count mutation (e.g. C02 edit), the
    service emits one ``ChangeRecord`` per changed field.
    """

    id: UUID
    nombre_tabla: str
    id_expediente: UUID
    nombre_campo: str | None
    valor_inicial: str | None
    valor_final: str | None
    fecha_cambio: datetime
    id_usuario_cambio: UUID
    accion: str  # "alta" | "edit" | "delete" | "transition"

    def __post_init__(self) -> None:
        if not self.nombre_tabla:
            raise ValueError("nombre_tabla is required")
        if not self.accion:
            raise ValueError("accion is required")
        if self.id_expediente is None:
            raise ValueError("id_expediente is required (D-EXP-3 deny-by-default)")
        if self.id_usuario_cambio is None:
            raise ValueError("id_usuario_cambio is required (D-EXP-3 deny-by-default)")
        if (
            self.fecha_cambio.tzinfo is None
            or self.fecha_cambio.tzinfo.utcoffset(self.fecha_cambio) is None
        ):
            raise ValueError(
                "fecha_cambio must be timezone-aware (got naive); "
                "audit rows must be reproducible across servers"
            )


class AuditLogPort(Protocol):
    """Append-only audit-log boundary."""

    async def append(self, event: ExpedienteAuditEvent) -> None: ...
    async def list_for_actor(
        self, actor_id: UUID, since: datetime
    ) -> Sequence[ExpedienteAuditEvent]: ...
    async def record_change(self, change: ChangeRecord) -> None:
        """Persist a ``ChangeRecord`` row inside the active transaction.

        The caller is expected to be inside a UoW block (DA-11). The
        implementation MUST NOT commit independently: it relies on the
        UoW for atomicity with the cabecera and the read-model.
        """
        ...


__all__ = ["AuditLogPort", "ChangeRecord", "ExpedienteAuditEvent"]
