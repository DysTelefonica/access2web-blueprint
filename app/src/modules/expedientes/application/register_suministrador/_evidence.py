"""Audit evidence builders for EXP-CAP-014."""

from datetime import datetime
from uuid import uuid4

from app.src.modules.expedientes.application.register_suministrador.command import (
    RegisterSuministradorCommand,
)
from app.src.modules.expedientes.ports.audit_log import ChangeRecord, ExpedienteAuditEvent


def audit_event(command: RegisterSuministradorCommand, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="suministrador.registered",
        actor_id=command.actor_id,
        target_id=command.expediente_id,
        capacidad="EXP-CAP-014",
        created_at=at,
        payload={
            "suministrador_id": str(command.suministrador_id),
            "id_suministrador": str(command.id_suministrador),
            "id_padre": str(command.id_padre) if command.id_padre else None,
        },
    )


def change_record(command: RegisterSuministradorCommand, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_suministradores",
        id_expediente=command.expediente_id,
        nombre_campo=None,
        valor_inicial=None,
        valor_final=str(command.suministrador_id),
        fecha_cambio=at,
        id_usuario_cambio=command.actor_id,
        accion="alta",
    )


__all__ = ["audit_event", "change_record"]
