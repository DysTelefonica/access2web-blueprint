from collections.abc import Callable
from dataclasses import replace
from datetime import UTC, datetime
from typing import Any, cast
from uuid import uuid4

from app.src.modules.expedientes.application.autosave_general.command import (
    AutosaveGeneralAuthorizationError,
    AutosaveGeneralCommand,
    AutosaveGeneralConflictError,
    AutosaveGeneralError,
    AutosaveGeneralResult,
    AutosaveGeneralValidationError,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.ports.audit_log import (
    AuditLogPort,
    ChangeRecord,
    ExpedienteAuditEvent,
)
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort
from app.src.modules.expedientes.ports.idempotency import IdempotencyPort, IdempotencyRecord


class AutosaveGeneralService:
    def __init__(
        self,
        *,
        expediente_repo: ExpedienteRepositoryPort,
        idempotency: IdempotencyPort,
        audit_log: AuditLogPort,
        uow_factory: Callable[[], Any],
    ) -> None:
        self._repo = expediente_repo
        self._idempotency = idempotency
        self._audit = audit_log
        self._uow_factory = uow_factory

    async def execute(self, command: AutosaveGeneralCommand) -> AutosaveGeneralResult:
        self._validate(command)
        signature = command.signature()
        recorded = await self._idempotency.get(command.idempotency_key)
        if recorded is not None:
            if recorded.request_signature != signature:
                raise AutosaveGeneralConflictError("idempotency key reused with different payload")
            return cast(AutosaveGeneralResult, recorded.result)

        loaded = await self._repo.get_by_id(command.expediente_id)
        if loaded is None:
            raise AutosaveGeneralValidationError(f"expediente {command.expediente_id} not found")
        existing = cast(Expediente, loaded)
        if existing.version != command.expected_version:
            raise AutosaveGeneralConflictError(
                f"version stale: stored={existing.version}, incoming={command.expected_version}"
            )

        changed = self._changed_fields(existing, command)
        modified_at = datetime.now(UTC)
        edited = self._edited(existing, command, modified_at, bool(changed))
        result = AutosaveGeneralResult(
            expediente_id=edited.id,
            new_version=edited.version,
            modified_at=modified_at,
            fields_changed=tuple(name for name, _, _ in changed),
        )

        try:
            with self._uow_factory() as _session:
                if changed:
                    await self._repo.update(edited)
                    await self._audit.append(self._event(command))
                    for name, old, new in changed:
                        await self._audit.record_change(
                            self._change(command, name, old, new, modified_at)
                        )
                await self._idempotency.record(
                    IdempotencyRecord(command.idempotency_key, signature, result)
                )
        except Exception as exc:
            raise AutosaveGeneralError(
                f"autosave failed for {command.expediente_id}: {exc}"
            ) from exc
        return result

    @staticmethod
    def _validate(command: AutosaveGeneralCommand) -> None:
        if command.idempotency_key is None:
            raise AutosaveGeneralValidationError("idempotency_key is required")
        if command.expected_version < 1:
            raise AutosaveGeneralValidationError("expected_version must be >= 1")
        if command.actor_id is None:
            raise AutosaveGeneralAuthorizationError("actor_id is required")

    @staticmethod
    def _changed_fields(
        existing: Expediente, command: AutosaveGeneralCommand
    ) -> list[tuple[str, object, object]]:
        changed: list[tuple[str, object, object]] = []
        for name in ("fecha_inicio_contrato", "fecha_fin_contrato"):
            old, new = getattr(existing, name), getattr(command, name)
            if old != new:
                changed.append((name, old, new))
        return changed

    @staticmethod
    def _edited(
        existing: Expediente,
        command: AutosaveGeneralCommand,
        modified_at: datetime,
        has_changes: bool,
    ) -> Expediente:
        if not has_changes:
            return existing
        try:
            return replace(
                existing,
                fecha_inicio_contrato=command.fecha_inicio_contrato,
                fecha_fin_contrato=command.fecha_fin_contrato,
                version=existing.version + 1,
                updated_at=modified_at,
            )
        except ValueError as exc:
            raise AutosaveGeneralValidationError(str(exc)) from exc

    @staticmethod
    def _event(command: AutosaveGeneralCommand) -> ExpedienteAuditEvent:
        return ExpedienteAuditEvent(
            id=uuid4(),
            event_type="expediente.general.autosaved",
            actor_id=command.actor_id,
            target_id=command.expediente_id,
            capacidad="EXP-CAP-030",
        )

    @staticmethod
    def _change(
        command: AutosaveGeneralCommand,
        name: str,
        old: object,
        new: object,
        modified_at: datetime,
    ) -> ChangeRecord:
        return ChangeRecord(
            id=uuid4(),
            nombre_tabla="expedientes",
            id_expediente=command.expediente_id,
            nombre_campo=name,
            valor_inicial=str(old) if old is not None else None,
            valor_final=str(new) if new is not None else None,
            fecha_cambio=modified_at,
            id_usuario_cambio=command.actor_id,
            accion="edit",
        )
