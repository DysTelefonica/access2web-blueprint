"""Transactional EXP-CAP-009 modification history registration."""

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast

from app.src.modules.expedientes.application.register_modificado._evidence import (
    audit_event,
    change_record,
)
from app.src.modules.expedientes.application.register_modificado.command import (
    RegisterModificadoAuthorizationError,
    RegisterModificadoCommand,
    RegisterModificadoConflictError,
    RegisterModificadoError,
    RegisterModificadoResult,
    RegisterModificadoValidationError,
)
from app.src.modules.expedientes.domain.modificado import Modificado
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort
from app.src.modules.expedientes.ports.modificado_repository import ModificadoRepositoryPort


class RegisterModificadoService:
    """Register one modification entry and its evidence in the same UoW."""

    def __init__(
        self,
        *,
        expediente_repo: ExpedienteRepositoryPort,
        modificado_repo: ModificadoRepositoryPort,
        audit_log: AuditLogPort,
        uow_factory: Callable[[], Any],
    ) -> None:
        self._expediente_repo = expediente_repo
        self._modificado_repo = modificado_repo
        self._audit_log = audit_log
        self._uow_factory = uow_factory

    async def execute(self, command: RegisterModificadoCommand) -> RegisterModificadoResult:
        self._validate(command)
        existing = await self._load_existing(command)
        if existing is not None:
            if not self._same_payload(existing, command):
                raise RegisterModificadoConflictError(
                    f"modificado {command.modificado_id!r} already exists with different payload"
                )
            return RegisterModificadoResult(
                existing.id, existing.id_expediente, existing.created_at
            )

        try:
            loaded = await self._expediente_repo.get_by_id(command.expediente_id)
        except Exception as exc:
            raise RegisterModificadoError("expediente dependency unavailable") from exc
        if loaded is None:
            raise RegisterModificadoValidationError(
                f"expediente {command.expediente_id!r} not found"
            )
        registered_at = datetime.now(UTC)
        try:
            modificado = Modificado(
                id=command.modificado_id,
                id_expediente=command.expediente_id,
                n_modificado=command.n_modificado,
                fecha_firma=command.fecha_firma,
                fecha_fin=command.fecha_fin,
                descripcion=command.descripcion,
                created_at=registered_at,
            )
        except ValueError as exc:
            raise RegisterModificadoValidationError(str(exc)) from exc

        try:
            with self._uow_factory():
                await self._modificado_repo.upsert(modificado)
                await self._audit_log.append(audit_event(command, registered_at))
                await self._audit_log.record_change(change_record(command, registered_at))
        except RegisterModificadoError:
            raise
        except Exception as exc:
            raise RegisterModificadoError(
                f"modificado registration failed for {command.expediente_id!r}: {exc}"
            ) from exc

        return RegisterModificadoResult(command.modificado_id, command.expediente_id, registered_at)

    async def _load_existing(self, command: RegisterModificadoCommand) -> Modificado | None:
        try:
            return cast(
                Modificado | None, await self._modificado_repo.get_by_id(command.modificado_id)
            )
        except Exception as exc:
            raise RegisterModificadoError("modificado dependency unavailable") from exc

    @staticmethod
    def _validate(command: RegisterModificadoCommand) -> None:
        if command.actor_id is None:
            raise RegisterModificadoAuthorizationError("actor_id is required (deny-by-default)")
        if command.modificado_id is None or command.expediente_id is None:
            raise RegisterModificadoValidationError("modificado_id and expediente_id are required")

    @staticmethod
    def _same_payload(existing: Modificado, command: RegisterModificadoCommand) -> bool:
        return (
            existing.id_expediente,
            existing.n_modificado,
            existing.fecha_firma,
            existing.fecha_fin,
            existing.descripcion,
        ) == (
            command.expediente_id,
            command.n_modificado,
            command.fecha_firma,
            command.fecha_fin,
            command.descripcion,
        )
