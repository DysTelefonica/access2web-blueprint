"""Audit evidence builders for EXP-CAP-010."""

from datetime import datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.anexos.command import (
    CreateAnexoCommand,
    DeleteAnexoCommand,
)
from app.src.modules.expedientes.domain.anexo.agregado import Anexo
from app.src.modules.expedientes.ports.audit_log import ChangeRecord, ExpedienteAuditEvent


def create_event(command: CreateAnexoCommand, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="anexo.created",
        actor_id=command.actor_id,
        target_id=command.expediente_id,
        capacidad="EXP-CAP-010",
        created_at=at,
        payload={
            "anexo_id": str(command.anexo_id),
            "storage_ref": command.referencia.storage_ref,
        },
    )


def create_change(command: CreateAnexoCommand, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_anexos",
        id_expediente=command.expediente_id,
        nombre_campo=None,
        valor_inicial=None,
        valor_final=str(command.anexo_id),
        fecha_cambio=at,
        id_usuario_cambio=command.actor_id,
        accion="alta",
    )


def delete_event(command: DeleteAnexoCommand, anexo: Anexo, at: datetime) -> ExpedienteAuditEvent:
    assert command.actor_id is not None
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="anexo.deleted",
        actor_id=command.actor_id,
        target_id=anexo.id_expediente,
        capacidad="EXP-CAP-010",
        created_at=at,
        payload={"anexo_id": str(anexo.id)},
    )


def delete_change(command: DeleteAnexoCommand, anexo: Anexo, at: datetime) -> ChangeRecord:
    actor_id: UUID = _require_actor(command.actor_id)
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_anexos",
        id_expediente=anexo.id_expediente,
        nombre_campo=None,
        valor_inicial=str(anexo.id),
        valor_final=None,
        fecha_cambio=at,
        id_usuario_cambio=actor_id,
        accion="delete",
    )


def _require_actor(actor_id: UUID | None) -> UUID:
    if actor_id is None:
        raise DeleteAnexoActorMissingError("actor_id is required")
    return actor_id


class DeleteAnexoActorMissingError(ValueError):
    """Internal guard for the audit evidence builders."""


__all__ = [
    "create_change",
    "create_event",
    "delete_change",
    "delete_event",
]
