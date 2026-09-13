"""Transactional EXP-CAP-008 milestone registration."""

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast

from app.src.modules.expedientes.application.register_hito._evidence import (
    audit_event,
    change_record,
)
from app.src.modules.expedientes.application.register_hito.command import (
    RegisterHitoAuthorizationError,
    RegisterHitoCommand,
    RegisterHitoConflictError,
    RegisterHitoError,
    RegisterHitoResult,
    RegisterHitoValidationError,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.hito import Hito
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort
from app.src.modules.expedientes.ports.hito_repository import HitoRepositoryPort


class RegisterHitoService:
    """Register one milestone and its evidence in the same unit of work."""

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

    async def execute(self, command: RegisterHitoCommand) -> RegisterHitoResult:
        self._validate(command)
        existing = await self._load_existing(command)
        if existing is not None:
            if not self._same_payload(existing, command):
                raise RegisterHitoConflictError(
                    f"hito {command.hito_id!r} already exists with different payload"
                )
            return RegisterHitoResult(existing.id, existing.id_expediente, existing.created_at)

        try:
            loaded = await self._expediente_repo.get_by_id(command.expediente_id)
        except Exception as exc:
            raise RegisterHitoError("expediente dependency unavailable") from exc
        if loaded is None:
            raise RegisterHitoValidationError(f"expediente {command.expediente_id!r} not found")
        expediente = cast(Expediente, loaded)
        registered_at = datetime.now(UTC)
        try:
            hito = Hito(
                id=command.hito_id,
                id_expediente=command.expediente_id,
                descripcion=command.descripcion,
                fecha_hito=command.fecha_hito,
                garantia_fecha_fin=command.garantia_fecha_fin,
                importe=command.importe,
                estado=expediente.estado,
                created_at=registered_at,
            )
        except ValueError as exc:
            raise RegisterHitoValidationError(str(exc)) from exc

        try:
            with self._uow_factory():
                await self._hito_repo.upsert(hito)
                await self._audit_log.append(audit_event(command, registered_at))
                await self._audit_log.record_change(change_record(command, registered_at))
        except RegisterHitoError:
            raise
        except Exception as exc:
            raise RegisterHitoError(
                f"milestone registration failed for {command.expediente_id!r}: {exc}"
            ) from exc

        return RegisterHitoResult(command.hito_id, command.expediente_id, registered_at)

    async def _load_existing(self, command: RegisterHitoCommand) -> Hito | None:
        try:
            return cast(Hito | None, await self._hito_repo.get_by_id(command.hito_id))
        except Exception as exc:
            raise RegisterHitoError("milestone dependency unavailable") from exc

    @staticmethod
    def _validate(command: RegisterHitoCommand) -> None:
        if command.actor_id is None:
            raise RegisterHitoAuthorizationError("actor_id is required (deny-by-default)")
        if command.hito_id is None or command.expediente_id is None:
            raise RegisterHitoValidationError("hito_id and expediente_id are required")

    @staticmethod
    def _same_payload(existing: Hito, command: RegisterHitoCommand) -> bool:
        return (
            existing.id_expediente,
            existing.descripcion,
            existing.fecha_hito,
            existing.garantia_fecha_fin,
            existing.importe,
        ) == (
            command.expediente_id,
            command.descripcion,
            command.fecha_hito,
            command.garantia_fecha_fin,
            command.importe,
        )
