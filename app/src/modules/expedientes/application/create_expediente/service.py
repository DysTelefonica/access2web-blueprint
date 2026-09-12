"""ExpedienteAltaService — CAP-001 alta transaccional (issue #226).

CAP-001 §Camino feliz: registra cabecera, hijos, read-model y último
cambio en una transacción confirmada. El servicio encapsula la
lógica de aplicación — el delivery layer (HTTP routes, CLI driver)
la invoca; la capa de persistencia (UoW + repositories) le da la
transaccionalidad.

DA-1: la capa de aplicación NO toca SQLAlchemy directamente. Pasa a
través de la sesión del UoW (F04) y de los puertos (F01).

Notas de implementación:

- La "vista materializada" (``TbExpedientesConEntidades``, D113) no
  se regenera aquí — eso es vertical separado (R/CAP-013/014). CAP-001
  pide "read-model" en general; el primer paso materializado es la
  cabecera + un hito inicial + un audit event.
- El "último cambio" (``TbUltimoCambio``) tampoco se materializa aquí
  — el audit event cubre la trazabilidad mínima. El vertical E2E
  (E01+) puede enriquecer la tabla cuando llegue.
- El AM (sin padre) crea un hito inicial "Inicio del expediente" como
  hijo obligatorio. LOTE y BASED requieren id_expediente_padre, que
  ya valida ``Expediente.__post_init__``; el servicio captura el
  ValueError y lo convierte a ``ExpedienteAltaValidationError``.
"""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.create_expediente.command import (
    ExpedienteAltaAuthorizationError,
    ExpedienteAltaCommand,
    ExpedienteAltaError,
    ExpedienteAltaResult,
    ExpedienteAltaValidationError,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.hito import Hito
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.expediente_repository import (
    ExpedienteRepositoryPort,
)
from app.src.modules.expedientes.ports.hito_repository import HitoRepositoryPort


class ExpedienteAltaService:
    """Use case: create an expediente (CAP-001).

    Validation happens before the UoW opens (so a malformed command
    doesn't waste a transaction). Persistence happens inside the UoW
    so a repository failure rolls everything back. Audit happens
    inside the same transaction (DA-11 audit-in-same-tx).
    """

    def __init__(
        self,
        *,
        expediente_repo: ExpedienteRepositoryPort,
        hito_repo: HitoRepositoryPort,
        audit_log: AuditLogPort,
        uow_factory: Callable[[], Any],
    ) -> None:
        self._expediente_repo = expediente_repo
        self._hito_repo = hito_repo
        self._audit_log = audit_log
        self._uow_factory = uow_factory

    async def execute(self, command: ExpedienteAltaCommand) -> ExpedienteAltaResult:
        """Run the alta end-to-end.

        Raises:
            ExpedienteAltaValidationError: malformed input or
                domain invariant violated (LOTE without padre, etc.).
            ExpedienteAltaAuthorizationError: actor lacks permission.
            ExpedienteAltaConflictError: codigo_ordinal already taken.
            ExpedienteAltaError: any other persistence failure.
        """
        self._validate(command)

        expediente_id = uuid4()
        version = 1
        created_at = datetime.now(UTC)

        try:
            aggregate = Expediente(
                id=expediente_id,
                tipo=command.tipo,
                estado=ExpedienteEstado.BORRADOR,
                version=version,
                id_expediente_padre=command.id_expediente_padre,
                created_at=created_at,
            )
        except ValueError as exc:
            # ``Expediente.__post_init__`` enforces LOTE/BASED need a padre
            # and version >= 1. Surface those as domain validation errors.
            raise ExpedienteAltaValidationError(str(exc)) from exc

        initial_hito = Hito(
            id=uuid4(),
            id_expediente=expediente_id,
            fecha_hito=created_at.date(),
            garantia_fecha_fin=None,
            estado=ExpedienteEstado.BORRADOR,
        )

        try:
            uow = self._uow_factory()
            with uow as _session:
                await self._expediente_repo.create(aggregate)
                await self._hito_repo.upsert(initial_hito)
                await self._audit_log.append(
                    self._audit_event(
                        command=command,
                        expediente_id=expediente_id,
                        action="expediente.created",
                        result="ok",
                    )
                )
        except ExpedienteAltaError:
            # Re-raise domain errors verbatim; only the UoW roll-back path
            # is the use case's responsibility.
            raise
        except Exception as exc:
            # Persisted failure: wrap as a domain error so the caller can
            # react without inspecting the underlying exception class.
            raise ExpedienteAltaError(
                f"alta failed for ordinal {command.codigo_ordinal!r}: {exc}"
            ) from exc

        return ExpedienteAltaResult(
            expediente_id=expediente_id,
            version=version,
            created_at=created_at,
        )

    @staticmethod
    def _validate(command: ExpedienteAltaCommand) -> None:
        """Check domain invariants before opening the UoW.

        Empty / whitespace ordinals are CAP-001 §Validación failures
        (400). The LOTE/BASED-without-padre check is delegated to
        ``Expediente.__post_init__`` (F02) and caught here as a domain
        validation error.
        """
        if not command.codigo_ordinal or not command.codigo_ordinal.strip():
            raise ExpedienteAltaValidationError(
                "codigo_ordinal is required and cannot be whitespace-only"
            )
        if command.actor_id is None:
            raise ExpedienteAltaAuthorizationError("actor_id is required (D-EXP-3 deny-by-default)")

    @staticmethod
    def _audit_event(
        *,
        command: ExpedienteAltaCommand,
        expediente_id: UUID,
        action: str,
        result: str,
    ) -> Any:
        """Build the audit event the audit log expects.

        Mirrors ``ExpedienteAuditEvent`` (F01 ports/audit_log.py) but
        keeps this module free of that import so the test fakes can
        inject their own AuditLogPort implementation.
        """
        from app.src.modules.expedientes.ports.audit_log import (
            ExpedienteAuditEvent,
        )

        return ExpedienteAuditEvent(
            id=uuid4(),
            event_type=action,
            actor_id=command.actor_id,
            target_id=expediente_id,
            capacidad="EXP-CAP-001",
            module="expedientes",
            result=result,
        )
