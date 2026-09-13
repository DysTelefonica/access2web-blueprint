"""ExpedienteTransitionService — CAP-004 cambio de tipo (issue #229).

CAP-004 §Camino feliz: "se aplica el nuevo tipo y sus efectos derivados
de forma consistente, con salida determinista y auditable". The
allowed-transitions matrix encodes the business rules.

§Concurrencia o fallo: "preserva invariantes, no duplica ni pierde
evidencia, queda reintentable".

The implementation splits ``execute()`` into private helpers so
the complexity ceiling (10) holds. Each helper is a single
responsibility; ``execute()`` reads as a pipeline.

DA-1: the application layer does NOT touch SQLAlchemy directly.
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any, Callable, cast
from uuid import UUID, uuid4

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
        """Run the transition end-to-end.

        Reads as a 5-step pipeline: validate → load+version → resolve
        → check matrix → persist. Each step is a private helper so
        the function's complexity stays under the 10-statement
        ceiling enforced by check_complexity.py.
        """
        self._validate(command)

        existing = await self._load_aggregate(command.expediente_id)
        self._check_version(existing, command.expected_version)

        new_tipo, new_padre = self._resolve_target_fields(command, existing)
        if self._is_noop(existing, new_tipo, new_padre):
            return self._noop_result(existing, new_tipo, new_padre)

        self._check_transition_allowed(existing.tipo, new_tipo)
        edited = self._build_edited_aggregate(existing, new_tipo, new_padre)
        modified_at = datetime.now(UTC)
        return await self._persist_within_uow(
            command=command,
            existing=existing,
            edited=edited,
            modified_at=modified_at,
        )

    async def _load_aggregate(self, expediente_id: Any) -> Expediente:
        loaded = await self._expediente_repo.get_by_id(expediente_id)
        if loaded is None:
            raise ExpedienteTransitionValidationError(f"expediente {expediente_id!r} not found")
        return cast(Expediente, loaded)

    @staticmethod
    def _check_version(existing: Expediente, expected_version: int) -> None:
        if existing.version != expected_version:
            raise ExpedienteTransitionConflictError(
                f"version stale: stored={existing.version}, incoming={expected_version}"
            )

    @staticmethod
    def _resolve_target_fields(
        command: ExpedienteTransitionCommand, existing: Expediente
    ) -> tuple[Any, Any]:
        """Return the new ``tipo`` and ``id_expediente_padre``.

        If the command did not supply a value, keep the existing one.
        """
        new_tipo = command.new_tipo if command.new_tipo is not None else existing.tipo
        new_padre = (
            command.new_id_expediente_padre
            if command.new_id_expediente_padre is not None
            else existing.id_expediente_padre
        )
        return new_tipo, new_padre

    @staticmethod
    def _is_noop(existing: Expediente, new_tipo: Any, new_padre: Any) -> bool:
        return new_tipo == existing.tipo and new_padre == existing.id_expediente_padre

    @staticmethod
    def _noop_result(
        existing: Expediente, new_tipo: Any, new_padre: Any
    ) -> ExpedienteTransitionResult:
        return ExpedienteTransitionResult(
            expediente_id=existing.id,
            new_tipo=new_tipo,
            new_id_expediente_padre=new_padre,
            new_version=existing.version,
            modified_at=datetime.now(UTC),
        )

    @staticmethod
    def _check_transition_allowed(current_tipo: Any, new_tipo: Any) -> None:
        if new_tipo not in ALLOWED_TRANSITIONS.get(current_tipo, frozenset()):
            raise ExpedienteTransitionValidationError(
                f"transition {current_tipo} -> {new_tipo} not allowed"
            )

    @staticmethod
    def _build_edited_aggregate(existing: Expediente, new_tipo: Any, new_padre: Any) -> Expediente:
        new_version = existing.version + 1
        try:
            return Expediente(
                id=existing.id,
                tipo=new_tipo,
                estado=existing.estado,
                version=new_version,
                id_expediente_padre=new_padre,
                created_at=existing.created_at,
                updated_at=datetime.now(UTC),
            )
        except ValueError as exc:
            raise ExpedienteTransitionValidationError(str(exc)) from exc

    async def _persist_within_uow(
        self,
        *,
        command: ExpedienteTransitionCommand,
        existing: Expediente,
        edited: Expediente,
        modified_at: datetime,
    ) -> ExpedienteTransitionResult:
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
                    new_tipo=edited.tipo,
                    new_padre=edited.id_expediente_padre,
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
        """
        records: list[ChangeRecord] = []
        if new_tipo != existing.tipo:
            records.append(
                ChangeRecord(
                    id=uuid4(),
                    nombre_tabla="expedientes",
                    id_expediente=command.expediente_id,
                    nombre_campo="tipo",
                    valor_inicial=str(existing.tipo),
                    valor_final=str(new_tipo),
                    fecha_cambio=modified_at,
                    id_usuario_cambio=command.actor_id,
                    accion="transition",
                )
            )
        if new_padre != existing.id_expediente_padre:
            records.append(
                ChangeRecord(
                    id=uuid4(),
                    nombre_tabla="expedientes",
                    id_expediente=command.expediente_id,
                    nombre_campo="id_expediente_padre",
                    valor_inicial=(
                        str(existing.id_expediente_padre)
                        if existing.id_expediente_padre is not None
                        else None
                    ),
                    valor_final=str(new_padre) if new_padre is not None else None,
                    fecha_cambio=modified_at,
                    id_usuario_cambio=command.actor_id,
                    accion="transition",
                )
            )
        return records
