"""Audit evidence builders for EXP-CAP-012."""

from datetime import datetime
from uuid import uuid4

from app.src.modules.expedientes.application.register_responsable.command import (
    RegisterResponsableCommand,
)
from app.src.modules.expedientes.ports.audit_log import ChangeRecord, ExpedienteAuditEvent


def audit_event(command: RegisterResponsableCommand, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="responsable.registered",
        actor_id=command.actor_id,
        target_id=command.expediente_id,
        capacidad="EXP-CAP-012",
        created_at=at,
        payload={
            "responsable_id": str(command.responsable_id),
            "usuario_id": str(command.usuario_id),
            "rol": command.rol,
        },
    )


def change_record(command: RegisterResponsableCommand, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_responsables",
        id_expediente=command.expediente_id,
        nombre_campo=None,
        valor_inicial=None,
        valor_final=str(command.responsable_id),
        fecha_cambio=at,
        id_usuario_cambio=command.actor_id,
        accion="alta",
    )


__all__ = ["audit_event", "change_record"]
