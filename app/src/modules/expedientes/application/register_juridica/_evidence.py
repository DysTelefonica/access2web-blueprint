"""Audit evidence builders for EXP-CAP-013."""

from datetime import datetime
from uuid import uuid4

from app.src.modules.expedientes.application.register_juridica.command import (
    RegisterJuridicaCommand,
)
from app.src.modules.expedientes.ports.audit_log import ChangeRecord, ExpedienteAuditEvent


def audit_event(command: RegisterJuridicaCommand, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="juridica.registered",
        actor_id=command.actor_id,
        target_id=command.expediente_id,
        capacidad="EXP-CAP-013",
        created_at=at,
        payload={
            "juridica_id": str(command.juridica_id),
            "id_juridica": str(command.id_juridica),
        },
    )


def change_record(command: RegisterJuridicaCommand, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_juridicas",
        id_expediente=command.expediente_id,
        nombre_campo=None,
        valor_inicial=None,
        valor_final=str(command.juridica_id),
        fecha_cambio=at,
        id_usuario_cambio=command.actor_id,
        accion="alta",
    )


__all__ = ["audit_event", "change_record"]
