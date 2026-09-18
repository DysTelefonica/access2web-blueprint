from datetime import datetime
from uuid import uuid4

from app.src.modules.expedientes.application.register_hito.command import RegisterHitoCommand
from app.src.modules.expedientes.ports.audit_log import ChangeRecord, ExpedienteAuditEvent


def audit_event(command: RegisterHitoCommand, at: datetime) -> ExpedienteAuditEvent:
    return ExpedienteAuditEvent(
        id=uuid4(),
        event_type="hito.registered",
        actor_id=command.actor_id,
        target_id=command.expediente_id,
        capacidad="EXP-CAP-008",
        created_at=at,
        payload={"hito_id": str(command.hito_id)},
    )


def change_record(command: RegisterHitoCommand, at: datetime) -> ChangeRecord:
    return ChangeRecord(
        id=uuid4(),
        nombre_tabla="expedientes_hitos",
        id_expediente=command.expediente_id,
        nombre_campo=None,
        valor_inicial=None,
        valor_final=str(command.hito_id),
        fecha_cambio=at,
        id_usuario_cambio=command.actor_id,
        accion="alta",
    )
