"""Audit evidence builders for EXP-CAP-015..019 catalogs."""

from datetime import datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.catalogs.command import (
    CatalogQuery,
    CpvDecision,
    CpvQuery,
)
from app.src.modules.expedientes.ports.audit_log import (
    ChangeRecord,
    ExpedienteAuditEvent,
)


def query_event(query: CatalogQuery, at: datetime, *, capacidad: str) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="catalog.queried",
        actor_id=query.actor_id,
        target_id=query.actor_id,
        capacidad=capacidad,
        created_at=at,
        payload={"text": query.text, "limit": query.limit},
    )


def cpv_event(query: CpvQuery, decision: CpvDecision, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="cpv.decision",
        actor_id=query.actor_id,
        target_id=query.raw_code.digits and UUID(int=0) or UUID(int=0),
        capacidad="EXP-CAP-016",
        created_at=at,
        payload={
            "code": str(query.raw_code.digits),
            "decision": decision.decision,
            "accepted": decision.accepted,
        },
    )


def cpv_change(query: CpvQuery, decision: CpvDecision, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_cpv_decisions",
        id_expediente=UUID(int=0),
        nombre_campo="cpv_code",
        valor_inicial=None,
        valor_final=str(query.raw_code.digits),
        fecha_cambio=at,
        id_usuario_cambio=query.actor_id,
        accion="edit",
    )


__all__ = ["cpv_change", "cpv_event", "query_event"]
