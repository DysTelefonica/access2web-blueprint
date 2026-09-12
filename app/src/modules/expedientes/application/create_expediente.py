"""CreateExpediente — CAP-001, CAP-006, D-EXP-1, D-EXP-4.

Deny-by-default (D-EXP-3): rejections are auditable with result != 'ok'.
"""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.domain import Expediente, ExpedienteEstado, ExpedienteTipo
from app.src.modules.expedientes.ports.audit_log import AuditLogPort, ExpedienteAuditEvent
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort


class CreateExpedienteError(Exception):
    """Raised when an Expediente cannot be created."""

    def __init__(self, reason: str) -> None:
        self.reason = reason
        super().__init__(reason)


def _now_utc(now: datetime | None) -> datetime:
    if now is not None:
        return now
    return datetime.now(UTC)


async def create_expediente(
    *,
    tipo: ExpedienteTipo,
    estado: ExpedienteEstado = ExpedienteEstado.BORRADOR,
    id_expediente_padre: UUID | None = None,
    repository: ExpedienteRepositoryPort,
    audit: AuditLogPort,
    actor_id: UUID | None = None,
    correlation_id: UUID | None = None,
    now: datetime | None = None,
) -> Expediente:
    """Create a new Expediente aggregate root.

    Creates an AM (root), LOTE (requires AM or another Lote), or BASED
    (requires AM or Lote) Expediente. CAP-006: LOTE and BASED require
    id_expediente_padre.

    Side effects:
        1. ``repository.create(expediente)`` persists the aggregate.
        2. ``audit.append(event)`` records the operation.

    Raises:
        CreateExpedienteError: tipo requires parent but none was provided.

    Returns the canonical ``Expediente`` with a freshly-generated UUID.
    """
    ts = _now_utc(now)

    # Validate hierarchy invariant before construction
    if tipo.value in ("LOTE", "BASED") and id_expediente_padre is None:
        await _reject(
            event_type="expedientes.create",
            capacidad="EXP-CAP-001",
            actor_id=actor_id,
            correlation_id=correlation_id,
            audit=audit,
            reason=f"tipo={tipo.value} requires id_expediente_padre",
            now=ts,
        )
        raise CreateExpedienteError(f"tipo={tipo.value} requires id_expediente_padre")

    # Construct aggregate
    expediente = Expediente(
        id=uuid4(),
        tipo=tipo,
        estado=estado,
        version=1,
        id_expediente_padre=id_expediente_padre,
        created_at=ts,
        updated_at=ts,
    )

    # Persist
    await repository.create(expediente)

    # Audit success
    await audit.append(
        ExpedienteAuditEvent(
            id=uuid4(),
            event_type="expedientes.create",
            actor_id=actor_id or uuid4(),
            target_id=expediente.id,
            capacidad="EXP-CAP-001",
            module="expedientes",
            result="ok",
            correlation_id=correlation_id,
            payload={"tipo": tipo.value},
            created_at=ts,
        )
    )

    return expediente


async def _reject(
    event_type: str,
    capacidad: str,
    actor_id: UUID | None,
    correlation_id: UUID | None,
    audit: AuditLogPort,
    reason: str,
    now: datetime,
) -> None:
    """Append a denied audit event without persisting."""
    await audit.append(
        ExpedienteAuditEvent(
            id=uuid4(),
            event_type=event_type,
            actor_id=actor_id or uuid4(),
            target_id=uuid4(),
            capacidad=capacidad,
            module="expedientes",
            result="denied",
            correlation_id=correlation_id,
            payload={"reason": reason},
            created_at=now,
        )
    )
