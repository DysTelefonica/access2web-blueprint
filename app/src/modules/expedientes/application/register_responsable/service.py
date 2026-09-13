"""Transactional EXP-CAP-012 responsable assignment."""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast

from app.src.modules.expedientes.application.register_responsable._evidence import (
    audit_event,
    change_record,
)
from app.src.modules.expedientes.application.register_responsable.command import (
    RegisterResponsableAuthorizationError,
    RegisterResponsableCommand,
    RegisterResponsableConflictError,
    RegisterResponsableError,
    RegisterResponsableResult,
    RegisterResponsableValidationError,
)
from app.src.modules.expedientes.domain.responsable import Responsable
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort
from app.src.modules.expedientes.ports.responsable_repository import ResponsableRepositoryPort


class RegisterResponsableService:
    """Register one responsable assignment and its evidence in the same UoW."""

    def __init__(
        self,
        *,
        expediente_repo: ExpedienteRepositoryPort,
        responsable_repo: ResponsableRepositoryPort,
        audit_log: AuditLogPort,
        uow_factory: Callable[[], Any],
    ) -> None:
        self._expediente_repo = expediente_repo
        self._responsable_repo = responsable_repo
        self._audit_log = audit_log
        self._uow_factory = uow_factory

    async def execute(self, command: RegisterResponsableCommand) -> RegisterResponsableResult:
        self._validate(command)
        existing = await self._load_existing(command)
        if existing is not None:
            if not self._same_payload(existing, command):
                raise RegisterResponsableConflictError(
                    f"responsable {command.responsable_id!r} already exists with different payload"
                )
            return RegisterResponsableResult(
                existing.id, existing.id_expediente, existing.created_at or datetime.now(UTC)
            )

        try:
            loaded = await self._expediente_repo.get_by_id(command.expediente_id)
        except Exception as exc:
            raise RegisterResponsableError("expediente dependency unavailable") from exc
        if loaded is None:
            raise RegisterResponsableValidationError(
                f"expediente {command.expediente_id!r} not found"
            )
        registered_at = datetime.now(UTC)
        try:
            responsable = Responsable(
                id=command.responsable_id,
                id_expediente=command.expediente_id,
                id_usuario=command.usuario_id,
                rol=command.rol,
                correo_siempre=command.correo_siempre,
                es_jefe_proyecto=command.es_jefe_proyecto,
                es_preventa=command.es_preventa,
                created_at=registered_at,
            )
        except ValueError as exc:
            raise RegisterResponsableValidationError(str(exc)) from exc

        try:
            with self._uow_factory():
                await self._responsable_repo.upsert(responsable)
                await self._audit_log.append(audit_event(command, registered_at))
                await self._audit_log.record_change(change_record(command, registered_at))
        except RegisterResponsableError:
            raise
        except Exception as exc:
            raise RegisterResponsableError(
                f"responsable registration failed for {command.expediente_id!r}: {exc}"
            ) from exc

        return RegisterResponsableResult(
            command.responsable_id, command.expediente_id, registered_at
        )

    async def _load_existing(self, command: RegisterResponsableCommand) -> Responsable | None:
        try:
            return cast(
                Responsable | None,
                await self._responsable_repo.get_by_id(command.responsable_id),
            )
        except Exception as exc:
            raise RegisterResponsableError("responsable dependency unavailable") from exc

    @staticmethod
    def _validate(command: RegisterResponsableCommand) -> None:
        if command.actor_id is None:
            raise RegisterResponsableAuthorizationError("actor_id is required (deny-by-default)")
        if (
            command.responsable_id is None
            or command.expediente_id is None
            or command.usuario_id is None
        ):
            raise RegisterResponsableValidationError(
                "responsable_id, expediente_id and usuario_id are required"
            )

    @staticmethod
    def _same_payload(existing: Responsable, command: RegisterResponsableCommand) -> bool:
        return (
            existing.id_expediente,
            existing.id_usuario,
            existing.rol,
            existing.correo_siempre,
            existing.es_jefe_proyecto,
            existing.es_preventa,
        ) == (
            command.expediente_id,
            command.usuario_id,
            command.rol,
            command.correo_siempre,
            command.es_jefe_proyecto,
            command.es_preventa,
        )
