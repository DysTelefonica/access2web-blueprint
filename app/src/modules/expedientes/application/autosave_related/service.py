"""Transactional EXP-CAP-031 autosave for related-data sections."""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast

from app.src.modules.expedientes.application.autosave_related._evidence import (
    anualidad_change,
    anualidad_event,
    hito_change,
    hito_event,
    modificado_change,
    modificado_event,
)
from app.src.modules.expedientes.application.autosave_related.command import (
    AutosaveRelatedAuthorizationError,
    AutosaveRelatedCommand,
    AutosaveRelatedConflictError,
    AutosaveRelatedError,
    AutosaveRelatedResult,
    AutosaveRelatedValidationError,
)
from app.src.modules.expedientes.domain.anualidad import Anualidad
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.hito import Hito
from app.src.modules.expedientes.domain.modificado import Modificado
from app.src.modules.expedientes.ports.anualidad_repository import AnualidadRepositoryPort
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort
from app.src.modules.expedientes.ports.hito_repository import HitoRepositoryPort
from app.src.modules.expedientes.ports.idempotency import IdempotencyPort, IdempotencyRecord
from app.src.modules.expedientes.ports.modificado_repository import ModificadoRepositoryPort


class AutosaveRelatedService:
    """Atomic autosave across hitos, modificados and anualidades."""

    def __init__(
        self,
        *,
        expediente_repo: ExpedienteRepositoryPort,
        hito_repo: HitoRepositoryPort,
        modificado_repo: ModificadoRepositoryPort,
        anualidad_repo: AnualidadRepositoryPort,
        idempotency: IdempotencyPort,
        audit_log: AuditLogPort,
        uow_factory: Callable[[], Any],
    ) -> None:
        self._expediente_repo = expediente_repo
        self._hito_repo = hito_repo
        self._modificado_repo = modificado_repo
        self._anualidad_repo = anualidad_repo
        self._idempotency = idempotency
        self._audit_log = audit_log
        self._uow_factory = uow_factory

    async def execute(self, command: AutosaveRelatedCommand) -> AutosaveRelatedResult:
        self._validate(command)
        replayed = await self._replay(command)
        if replayed is not None:
            return replayed
        expediente = await self._load_expediente(command.expediente_id)
        self._check_version(expediente, command.expected_version)
        registered_at = datetime.now(UTC)
        result = self._build_result(command, registered_at)
        try:
            with self._uow_factory():
                await self._persist(command, registered_at)
                await self._audit(command, registered_at)
                await self._idempotency.record(
                    IdempotencyRecord(command.idempotency_key, command.signature(), result)
                )
        except AutosaveRelatedError:
            raise
        except Exception as exc:
            raise AutosaveRelatedError(
                f"related autosave failed for {command.expediente_id!r}: {exc}"
            ) from exc
        return result

    async def _replay(self, command: AutosaveRelatedCommand) -> AutosaveRelatedResult | None:
        recorded = await self._idempotency.get(command.idempotency_key)
        if recorded is None:
            return None
        if recorded.request_signature != command.signature():
            raise AutosaveRelatedConflictError("idempotency key reused with different payload")
        return cast(AutosaveRelatedResult, recorded.result)

    async def _load_expediente(self, expediente_id: Any) -> Expediente:
        try:
            loaded = await self._expediente_repo.get_by_id(expediente_id)
        except Exception as exc:
            raise AutosaveRelatedError("expediente dependency unavailable") from exc
        if loaded is None:
            raise AutosaveRelatedValidationError(f"expediente {expediente_id!r} not found")
        return cast(Expediente, loaded)

    @staticmethod
    def _check_version(existing: Expediente, expected_version: int) -> None:
        if existing.version != expected_version:
            raise AutosaveRelatedConflictError(
                f"version stale: stored={existing.version}, incoming={expected_version}"
            )

    @staticmethod
    def _build_result(command: AutosaveRelatedCommand, at: datetime) -> AutosaveRelatedResult:
        return AutosaveRelatedResult(
            expediente_id=command.expediente_id,
            hitos_added=len(command.hitos),
            modificados_added=len(command.modificados),
            anualidades_added=len(command.anualidades),
            idempotency_key=command.idempotency_key,
            registered_at=at,
        )

    async def _persist(self, command: AutosaveRelatedCommand, _at: datetime) -> None:
        for hito in command.hitos:
            await self._hito_repo.upsert(hito)
        for mod in command.modificados:
            await self._modificado_repo.upsert(mod)
        for anu in command.anualidades:
            await self._anualidad_repo.upsert(anu)

    async def _audit(self, command: AutosaveRelatedCommand, at: datetime) -> None:
        if command.hitos:
            await self._audit_log.append(hito_event(command, at))
            await self._audit_log.record_change(hito_change(command, at))
        if command.modificados:
            await self._audit_log.append(modificado_event(command, at))
            await self._audit_log.record_change(modificado_change(command, at))
        if command.anualidades:
            await self._audit_log.append(anualidad_event(command, at))
            await self._audit_log.record_change(anualidad_change(command, at))

    @staticmethod
    def _validate(command: AutosaveRelatedCommand) -> None:
        if command.actor_id is None:
            raise AutosaveRelatedAuthorizationError("actor_id is required (deny-by-default)")
        if command.idempotency_key is None:
            raise AutosaveRelatedValidationError("idempotency_key is required")
        if command.expected_version < 1:
            raise AutosaveRelatedValidationError("expected_version must be >= 1")
        AutosaveRelatedService._validate_hitos(command)
        AutosaveRelatedService._validate_modificados(command)
        AutosaveRelatedService._validate_anualidades(command)

    @staticmethod
    def _validate_hitos(command: AutosaveRelatedCommand) -> None:
        for hito in command.hitos:
            try:
                Hito(
                    id=hito.id,
                    id_expediente=hito.id_expediente,
                    fecha_hito=hito.fecha_hito,
                    garantia_fecha_fin=hito.garantia_fecha_fin,
                    descripcion=hito.descripcion,
                    importe=hito.importe,
                    estado=hito.estado,
                )
            except ValueError as exc:
                raise AutosaveRelatedValidationError(str(exc)) from exc
            if hito.id_expediente != command.expediente_id:
                raise AutosaveRelatedValidationError("hito references a different expediente")

    @staticmethod
    def _validate_modificados(command: AutosaveRelatedCommand) -> None:
        for mod in command.modificados:
            try:
                Modificado(
                    id=mod.id,
                    id_expediente=mod.id_expediente,
                    n_modificado=mod.n_modificado,
                    fecha_firma=mod.fecha_firma,
                    fecha_fin=mod.fecha_fin,
                    descripcion=mod.descripcion,
                )
            except ValueError as exc:
                raise AutosaveRelatedValidationError(str(exc)) from exc
            if mod.id_expediente != command.expediente_id:
                raise AutosaveRelatedValidationError("modificado references a different expediente")

    @staticmethod
    def _validate_anualidades(command: AutosaveRelatedCommand) -> None:
        for anu in command.anualidades:
            try:
                Anualidad(
                    id=anu.id,
                    id_expediente=anu.id_expediente,
                    anio=anu.anio,
                    bi_iva=anu.bi_iva,
                    bi_ipsi=anu.bi_ipsi,
                    bi_igic=anu.bi_igic,
                    bi_exenta=anu.bi_exenta,
                    iva=anu.iva,
                    ipsi=anu.ipsi,
                    igic=anu.igic,
                    periodo_facturacion=anu.periodo_facturacion,
                )
            except ValueError as exc:
                raise AutosaveRelatedValidationError(str(exc)) from exc
            if anu.id_expediente != command.expediente_id:
                raise AutosaveRelatedValidationError("anualidad references a different expediente")
