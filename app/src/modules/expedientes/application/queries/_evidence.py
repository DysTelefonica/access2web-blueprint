"""Audit evidence builders for EXP-CAP-025..029 query services."""

from datetime import datetime
from uuid import uuid4

from app.src.modules.expedientes.application.queries.command import (
    BandejaQuery,
    ExportacionExcelCommand,
    ExportacionExcelJob,
    TareasCalculadasQuery,
    TareasResult,
)
from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent


def bandeja_event(query: BandejaQuery, at: datetime, *, total: int) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="bandeja.queried",
        actor_id=query.actor_id,
        target_id=query.actor_id,
        capacidad="EXP-CAP-025",
        created_at=at,
        payload={"limit": query.limit, "offset": query.offset, "total": total},
    )


def export_event(cmd: ExportacionExcelCommand, job: ExportacionExcelJob) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="export.excel.requested",
        actor_id=cmd.actor_id,
        target_id=cmd.actor_id,
        capacidad="EXP-CAP-028",
        created_at=job.requested_at,
        payload={
            "job_id": str(job.job_id),
            "formato": job.formato,
            "filtros": cmd.filtros,
        },
    )


def tareas_event(
    query: TareasCalculadasQuery, result: TareasResult, at: datetime
) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="tareas.calculated",
        actor_id=query.actor_id,
        target_id=query.actor_id,
        capacidad="EXP-CAP-029",
        created_at=at,
        payload={"total": result.total, "contadores": result.contadores},
    )


__all__ = ["bandeja_event", "export_event", "tareas_event"]
