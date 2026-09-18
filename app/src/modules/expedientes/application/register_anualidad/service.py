"""Transactional EXP-CAP-011 anualidad registration."""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast

from app.src.modules.expedientes.application.register_anualidad._evidence import (
    audit_event,
    change_record,
)
from app.src.modules.expedientes.application.register_anualidad.command import (
    RegisterAnualidadAuthorizationError,
    RegisterAnualidadCommand,
    RegisterAnualidadConflictError,
    RegisterAnualidadError,
    RegisterAnualidadResult,
    RegisterAnualidadValidationError,
)
from app.src.modules.expedientes.domain.anualidad import Anualidad
from app.src.modules.expedientes.ports.anualidad_repository import AnualidadRepositoryPort
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort


class RegisterAnualidadService:
    """Register one anualidad and its evidence in the same UoW."""

    def __init__(
        self,
        *,
        expediente_repo: ExpedienteRepositoryPort,
        anualidad_repo: AnualidadRepositoryPort,
        audit_log: AuditLogPort,
        uow_factory: Callable[[], Any],
    ) -> None:
        self._expediente_repo = expediente_repo
        self._anualidad_repo = anualidad_repo
        self._audit_log = audit_log
        self._uow_factory = uow_factory

    async def execute(self, command: RegisterAnualidadCommand) -> RegisterAnualidadResult:
        self._validate(command)
        existing = await self._load_existing(command)
        if existing is not None:
            if not self._same_payload(existing, command):
                raise RegisterAnualidadConflictError(
                    f"anualidad {command.anualidad_id!r} already exists with different payload"
                )
            return RegisterAnualidadResult(
                existing.id, existing.id_expediente, existing.created_at or datetime.now(UTC)
            )

        try:
            loaded = await self._expediente_repo.get_by_id(command.expediente_id)
        except Exception as exc:
            raise RegisterAnualidadError("expediente dependency unavailable") from exc
        if loaded is None:
            raise RegisterAnualidadValidationError(
                f"expediente {command.expediente_id!r} not found"
            )
        registered_at = datetime.now(UTC)
        try:
            anualidad = Anualidad(
                id=command.anualidad_id,
                id_expediente=command.expediente_id,
                anio=command.anio,
                bi_iva=command.bi_iva,
                bi_ipsi=command.bi_ipsi,
                bi_igic=command.bi_igic,
                bi_exenta=command.bi_exenta,
                iva=command.iva,
                ipsi=command.ipsi,
                igic=command.igic,
                periodo_facturacion=command.periodo_facturacion,
                created_at=registered_at,
            )
        except ValueError as exc:
            raise RegisterAnualidadValidationError(str(exc)) from exc

        try:
            with self._uow_factory():
                await self._anualidad_repo.upsert(anualidad)
                await self._audit_log.append(audit_event(command, registered_at))
                await self._audit_log.record_change(change_record(command, registered_at))
        except RegisterAnualidadError:
            raise
        except Exception as exc:
            raise RegisterAnualidadError(
                f"anualidad registration failed for {command.expediente_id!r}: {exc}"
            ) from exc

        return RegisterAnualidadResult(command.anualidad_id, command.expediente_id, registered_at)

    async def _load_existing(self, command: RegisterAnualidadCommand) -> Anualidad | None:
        try:
            return cast(
                Anualidad | None, await self._anualidad_repo.get_by_id(command.anualidad_id)
            )
        except Exception as exc:
            raise RegisterAnualidadError("anualidad dependency unavailable") from exc

    @staticmethod
    def _validate(command: RegisterAnualidadCommand) -> None:
        if command.actor_id is None:
            raise RegisterAnualidadAuthorizationError("actor_id is required (deny-by-default)")
        if command.anualidad_id is None or command.expediente_id is None:
            raise RegisterAnualidadValidationError("anualidad_id and expediente_id are required")

    @staticmethod
    def _same_payload(existing: Anualidad, command: RegisterAnualidadCommand) -> bool:
        return (
            existing.id_expediente,
            existing.anio,
            existing.bi_iva,
            existing.bi_ipsi,
            existing.bi_igic,
            existing.bi_exenta,
            existing.iva,
            existing.ipsi,
            existing.igic,
            existing.periodo_facturacion,
        ) == (
            command.expediente_id,
            command.anio,
            command.bi_iva,
            command.bi_ipsi,
            command.bi_igic,
            command.bi_exenta,
            command.iva,
            command.ipsi,
            command.igic,
            command.periodo_facturacion,
        )
