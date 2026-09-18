"""Audit evidence builders for EXP-CAP-009."""

from datetime import datetime
from uuid import uuid4

from app.src.modules.expedientes.application.register_modificado.command import (
    RegisterModificadoCommand,
)
from app.src.modules.expedientes.ports.audit_log import ChangeRecord, ExpedienteAuditEvent


def audit_event(command: RegisterModificadoCommand, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="modificado.registered",
        actor_id=command.actor_id,
        target_id=command.expediente_id,
        capacidad="EXP-CAP-009",
        created_at=at,
        payload={"modificado_id": str(command.modificado_id)},
    )


def change_record(command: RegisterModificadoCommand, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_modificados",
        id_expediente=command.expediente_id,
        nombre_campo=None,
        valor_inicial=None,
        valor_final=str(command.modificado_id),
        fecha_cambio=at,
        id_usuario_cambio=command.actor_id,
        accion="alta",
    )
