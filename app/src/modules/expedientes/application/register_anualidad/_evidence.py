"""Audit evidence builders for EXP-CAP-011."""

from datetime import datetime
from uuid import uuid4

from app.src.modules.expedientes.application.register_anualidad.command import (
    RegisterAnualidadCommand,
)
from app.src.modules.expedientes.ports.audit_log import ChangeRecord, ExpedienteAuditEvent


def audit_event(command: RegisterAnualidadCommand, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="anualidad.registered",
        actor_id=command.actor_id,
        target_id=command.expediente_id,
        capacidad="EXP-CAP-011",
        created_at=at,
        payload={"anualidad_id": str(command.anualidad_id), "anio": command.anio},
    )


def change_record(command: RegisterAnualidadCommand, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_anualidades",
        id_expediente=command.expediente_id,
        nombre_campo=None,
        valor_inicial=None,
        valor_final=str(command.anualidad_id),
        fecha_cambio=at,
        id_usuario_cambio=command.actor_id,
        accion="alta",
    )


__all__ = ["audit_event", "change_record"]
