"""Transactional EXP-CAP-013 juridica registration."""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast

from app.src.modules.expedientes.application.register_juridica._evidence import (
    audit_event,
    change_record,
)
from app.src.modules.expedientes.application.register_juridica.command import (
    RegisterJuridicaAuthorizationError,
    RegisterJuridicaCommand,
    RegisterJuridicaConflictError,
    RegisterJuridicaError,
    RegisterJuridicaResult,
    RegisterJuridicaValidationError,
)
from app.src.modules.expedientes.domain.juridica import Juridica
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort
from app.src.modules.expedientes.ports.juridica_repository import JuridicaRepositoryPort


class RegisterJuridicaService:
    """Register one juridica and its evidence in the same UoW."""

    def __init__(
        self,
        *,
        expediente_repo: ExpedienteRepositoryPort,
        juridica_repo: JuridicaRepositoryPort,
        audit_log: AuditLogPort,
        uow_factory: Callable[[], Any],
    ) -> None:
        self._expediente_repo = expediente_repo
        self._juridica_repo = juridica_repo
        self._audit_log = audit_log
        self._uow_factory = uow_factory

    async def execute(self, command: RegisterJuridicaCommand) -> RegisterJuridicaResult:
        self._validate(command)
        existing = await self._load_existing(command)
        if existing is not None:
            if not self._same_payload(existing, command):
                raise RegisterJuridicaConflictError(
                    f"juridica {command.juridica_id!r} already exists with different payload"
                )
            return RegisterJuridicaResult(
                existing.id, existing.id_expediente, existing.created_at or datetime.now(UTC)
            )

        try:
            loaded = await self._expediente_repo.get_by_id(command.expediente_id)
        except Exception as exc:
            raise RegisterJuridicaError("expediente dependency unavailable") from exc
        if loaded is None:
            raise RegisterJuridicaValidationError(f"expediente {command.expediente_id!r} not found")
        registered_at = datetime.now(UTC)
        try:
            juridica = Juridica(
                id=command.juridica_id,
                id_expediente=command.expediente_id,
                id_juridica=command.id_juridica,
                id_suministrador=command.id_suministrador,
                contratista_principal=command.contratista_principal,
                sub_contratista=command.sub_contratista,
                created_at=registered_at,
            )
        except ValueError as exc:
            raise RegisterJuridicaValidationError(str(exc)) from exc

        try:
            with self._uow_factory():
                await self._juridica_repo.upsert(juridica)
                await self._audit_log.append(audit_event(command, registered_at))
                await self._audit_log.record_change(change_record(command, registered_at))
        except RegisterJuridicaError:
            raise
        except Exception as exc:
            raise RegisterJuridicaError(
                f"juridica registration failed for {command.expediente_id!r}: {exc}"
            ) from exc

        return RegisterJuridicaResult(command.juridica_id, command.expediente_id, registered_at)

    async def _load_existing(self, command: RegisterJuridicaCommand) -> Juridica | None:
        try:
            return cast(Juridica | None, await self._juridica_repo.get_by_id(command.juridica_id))
        except Exception as exc:
            raise RegisterJuridicaError("juridica dependency unavailable") from exc

    @staticmethod
    def _validate(command: RegisterJuridicaCommand) -> None:
        if command.actor_id is None:
            raise RegisterJuridicaAuthorizationError("actor_id is required (deny-by-default)")
        if (
            command.juridica_id is None
            or command.expediente_id is None
            or command.id_juridica is None
        ):
            raise RegisterJuridicaValidationError(
                "juridica_id, expediente_id and id_juridica are required"
            )

    @staticmethod
    def _same_payload(existing: Juridica, command: RegisterJuridicaCommand) -> bool:
        return (
            existing.id_expediente,
            existing.id_juridica,
            existing.id_suministrador,
            existing.contratista_principal,
            existing.sub_contratista,
        ) == (
            command.expediente_id,
            command.id_juridica,
            command.id_suministrador,
            command.contratista_principal,
            command.sub_contratista,
        )
