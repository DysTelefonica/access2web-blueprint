"""Audit evidence builders for EXP-CAP-031 related autosave."""

from datetime import datetime
from uuid import uuid4

from app.src.modules.expedientes.application.autosave_related.command import (
    AutosaveRelatedCommand,
)
from app.src.modules.expedientes.ports.audit_log import ChangeRecord, ExpedienteAuditEvent


def hito_event(command: AutosaveRelatedCommand, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="hito.autosaved",
        actor_id=command.actor_id,
        target_id=command.expediente_id,
        capacidad="EXP-CAP-031",
        created_at=at,
        payload={"count": len(command.hitos)},
    )


def modificado_event(command: AutosaveRelatedCommand, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="modificado.autosaved",
        actor_id=command.actor_id,
        target_id=command.expediente_id,
        capacidad="EXP-CAP-031",
        created_at=at,
        payload={"count": len(command.modificados)},
    )


def anualidad_event(command: AutosaveRelatedCommand, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="anualidad.autosaved",
        actor_id=command.actor_id,
        target_id=command.expediente_id,
        capacidad="EXP-CAP-031",
        created_at=at,
        payload={"count": len(command.anualidades)},
    )


def hito_change(command: AutosaveRelatedCommand, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_hitos",
        id_expediente=command.expediente_id,
        nombre_campo=None,
        valor_inicial=None,
        valor_final=str(len(command.hitos)),
        fecha_cambio=at,
        id_usuario_cambio=command.actor_id,
        accion="edit",
    )


def modificado_change(command: AutosaveRelatedCommand, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_modificados",
        id_expediente=command.expediente_id,
        nombre_campo=None,
        valor_inicial=None,
        valor_final=str(len(command.modificados)),
        fecha_cambio=at,
        id_usuario_cambio=command.actor_id,
        accion="edit",
    )


def anualidad_change(command: AutosaveRelatedCommand, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_anualidades",
        id_expediente=command.expediente_id,
        nombre_campo=None,
        valor_inicial=None,
        valor_final=str(len(command.anualidades)),
        fecha_cambio=at,
        id_usuario_cambio=command.actor_id,
        accion="edit",
    )


__all__ = [
    "anualidad_change",
    "anualidad_event",
    "hito_change",
    "hito_event",
    "modificado_change",
    "modificado_event",
]
