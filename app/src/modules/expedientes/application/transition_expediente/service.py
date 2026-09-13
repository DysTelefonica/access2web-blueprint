"""ExpedienteTransitionService — CAP-004 cambio de tipo (issue #229).

CAP-004 §Camino feliz: "se aplica el nuevo tipo y sus efectos derivados
de forma consistente, con salida determinista y auditable". The
allowed-transitions matrix encodes the business rules:

- AM is the leaf type. AM↔AM (identity) is allowed.
- AM→LOTE/BASED requires a padre (D-EXP-4 / aggregate invariant).
- LOTE→AM, BASED→AM (demotion) are allowed when invariants hold.
- LOTE↔BASED is allowed (padre preserved).
- BASED→LOTE, LOTE→BASED are allowed (padre preserved).

This matrix is the source of truth. Any new transition rule lands
here; the use case enforces it. The service refuses invalid
transitions before opening the UoW (fail-fast).

§Concurrencia o fallo: "preserva invariantes, no duplica ni pierde
evidencia, queda reintentable". The version check + UoW rollback
cover this.

DA-1: the application layer does NOT touch SQLAlchemy directly.
"""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast
from uuid import uuid4

from app.src.modules.expedientes.application.transition_expediente.command import (
    ExpedienteTransitionAuthorizationError,
    ExpedienteTransitionCommand,
    ExpedienteTransitionConflictError,
    ExpedienteTransitionError,
    ExpedienteTransitionResult,
    ExpedienteTransitionValidationError,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo
from app.src.modules.expedientes.ports.audit_log import (
    AuditLogPort,
    ChangeRecord,
    ExpedienteAuditEvent,
)
from app.src.modules.expedientes.ports.expediente_repository import (
    ExpedienteRepositoryPort,
)
from app.src.modules.expedientes.ports.hito_repository import HitoRepositoryPort

# Transition matrix — the source of truth for valid type transitions.
# Conservative: only documented transitions; add new rules here.
ALLOWED_TRANSITIONS: dict[ExpedienteTipo, frozenset[ExpedienteTipo]] = {
    ExpedienteTipo.AM: frozenset({ExpedienteTipo.AM, ExpedienteTipo.LOTE, ExpedienteTipo.BASED}),
    ExpedienteTipo.LOTE: frozenset({ExpedienteTipo.AM, ExpedienteTipo.LOTE, ExpedienteTipo.BASED}),
    ExpedienteTipo.BASED: frozenset({ExpedienteTipo.AM, ExpedienteTipo.LOTE, ExpedienteTipo.BASED}),
}


class ExpedienteTransitionService:
    """Use case: transition an expediente's type (CAP-004)."""

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

    async def execute(self, command: ExpedienteTransitionCommand) -> ExpedienteTransitionResult:
        self._validate(command)

        loaded = await self._expediente_repo.get_by_id(command.expediente_id)
        if loaded is None:
            raise ExpedienteTransitionValidationError(
                f"expediente {command.expediente_id!r} not found"
            )
        existing: Expediente = cast(Expediente, loaded)

        if getattr(existing, "version", None) != command.expected_version:
            raise ExpedienteTransitionConflictError(
                f"version stale: stored={getattr(existing, 'version', None)}, "
                f"incoming={command.expected_version}"
            )

        new_tipo = (
            command.new_tipo if command.new_tipo is not None else getattr(existing, "tipo", None)
        )
        new_padre = (
            command.new_id_expediente_padre
            if command.new_id_expediente_padre is not None
            else getattr(existing, "id_expediente_padre", None)
        )

        # No-op: caller didn't change anything. Return success silently.
        if new_tipo == getattr(existing, "tipo", None) and new_padre == getattr(
            existing, "id_expediente_padre", None
        ):
            return ExpedienteTransitionResult(
                expediente_id=existing.id,
                new_tipo=new_tipo,
                new_id_expediente_padre=new_padre,
                new_version=existing.version,
                modified_at=datetime.now(UTC),
            )

        # Transition matrix check.
        current_tipo = getattr(existing, "tipo", None)
        if new_tipo not in ALLOWED_TRANSITIONS.get(current_tipo, frozenset()):
            raise ExpedienteTransitionValidationError(
                f"transition {current_tipo} -> {new_tipo} not allowed"
            )

        # Compute the new version. The aggregate's __post_init__ will
        # validate LOTE/BASED need a padre.
        new_version = getattr(existing, "version", None) + 1
        try:
            edited = Expediente(
                id=getattr(existing, "id", None),
                tipo=new_tipo,
                estado=getattr(existing, "estado", None),
                version=new_version,
                id_expediente_padre=new_padre,
                created_at=getattr(existing, "created_at", None),
                updated_at=datetime.now(UTC),
            )
        except ValueError as exc:
            raise ExpedienteTransitionValidationError(str(exc)) from exc

        modified_at = datetime.now(UTC)

        try:
            uow = self._uow_factory()
            with uow as _session:
                await self._expediente_repo.update(edited)
                await self._audit_log.append(
                    self._audit_event(
                        command=command,
                        existing=existing,
                        modified_at=modified_at,
                    )
                )
                for change_record in self._change_records(
                    command=command,
                    existing=existing,
                    new_tipo=new_tipo,
                    new_padre=new_padre,
                    modified_at=modified_at,
                ):
                    await self._audit_log.record_change(change_record)
        except ExpedienteTransitionError:
            raise
        except Exception as exc:
            raise ExpedienteTransitionError(
                f"transition failed for {command.expediente_id!r}: {exc}"
            ) from exc

        return ExpedienteTransitionResult(
            expediente_id=edited.id,
            new_tipo=edited.tipo,
            new_id_expediente_padre=edited.id_expediente_padre,
            new_version=edited.version,
            modified_at=modified_at,
        )

    @staticmethod
    def _validate(command: ExpedienteTransitionCommand) -> None:
        if command.expediente_id is None:
            raise ExpedienteTransitionValidationError("expediente_id is required")
        if command.expected_version is None or command.expected_version < 1:
            raise ExpedienteTransitionValidationError(
                "expected_version is required and must be >= 1 (D-EXP-4)"
            )
        if command.actor_id is None:
            raise ExpedienteTransitionAuthorizationError(
                "actor_id is required (D-EXP-3 deny-by-default)"
            )

    @staticmethod
    def _audit_event(
        *,
        command: ExpedienteTransitionCommand,
        existing: Expediente,
        modified_at: datetime,
    ) -> ExpedienteAuditEvent:
        return ExpedienteAuditEvent(
            id=uuid4(),
            event_type="expediente.transitioned",
            actor_id=command.actor_id,
            target_id=command.expediente_id,
            capacidad="EXP-CAP-004",
            module="expedientes",
            result="ok",
        )

    @staticmethod
    def _change_records(
        *,
        command: ExpedienteTransitionCommand,
        existing: Expediente,
        new_tipo: Any,
        new_padre: Any,
        modified_at: datetime,
    ) -> list[ChangeRecord]:
        """One ``ChangeRecord`` per modified field.

        The cambio de tipo typically changes both ``tipo`` and
        ``id_expediente_padre`` (when transitioning to/from LOTE/BASED
        that requires a padre). Each changed field is its own
        ``ChangeRecord`` row, matching the C02 pattern.

        Only the fields that actually changed get a row. The caller
        sees the audit trace complete with one row per field.
        """
        records: list[ChangeRecord] = []
        if new_tipo != getattr(existing, "tipo", None):
            records.append(
                ChangeRecord(
                    id=uuid4(),
                    nombre_tabla="expedientes",
                    id_expediente=command.expediente_id,
                    nombre_campo="tipo",
                    valor_inicial=str(getattr(existing, "tipo", None)),
                    valor_final=str(new_tipo),
                    fecha_cambio=modified_at,
                    id_usuario_cambio=command.actor_id,
                    accion="transition",
                )
            )
        if new_padre != getattr(existing, "id_expediente_padre", None):
            records.append(
                ChangeRecord(
                    id=uuid4(),
                    nombre_tabla="expedientes",
                    id_expediente=command.expediente_id,
                    nombre_campo="id_expediente_padre",
                    valor_inicial=(
                        str(getattr(existing, "id_expediente_padre", None))
                        if getattr(existing, "id_expediente_padre", None) is not None
                        else None
                    ),
                    valor_final=str(new_padre) if new_padre is not None else None,
                    fecha_cambio=modified_at,
                    id_usuario_cambio=command.actor_id,
                    accion="transition",
                )
            )
        return records
