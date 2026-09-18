"""Transactional EXP-CAP-010 anexos service."""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast

from app.src.modules.expedientes.application.anexos._evidence import (
    create_change,
    create_event,
    delete_change,
    delete_event,
)
from app.src.modules.expedientes.application.anexos.command import (
    CreateAnexoAuthorizationError,
    CreateAnexoCommand,
    DeleteAnexoAuthorizationError,
    DeleteAnexoCommand,
    DeleteAnexoError,
    DeleteAnexoResult,
    RegisterAnexoError,
    RegisterAnexoResult,
    RegisterAnexoSizeLimitError,
    RegisterAnexoValidationError,
)
from app.src.modules.expedientes.domain.anexo.agregado import Anexo, AnexoEvent
from app.src.modules.expedientes.domain.anexo.retention import RetentionExpired
from app.src.modules.expedientes.domain.anexo.size_limit import (
    SizeLimitExceeded,
    SizeLimitPolicy,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.ports.anexo_repository import AnexoRepositoryPort
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort


class AnexosService:
    """Create and delete anexos inside one UoW each (CAP-010)."""

    def __init__(
        self,
        *,
        expediente_repo: ExpedienteRepositoryPort,
        anexo_repo: AnexoRepositoryPort,
        audit_log: AuditLogPort,
        size_limit: SizeLimitPolicy,
        uow_factory: Callable[[], Any],
    ) -> None:
        self._expediente_repo = expediente_repo
        self._anexo_repo = anexo_repo
        self._audit_log = audit_log
        self._size_limit = size_limit
        self._uow_factory = uow_factory

    async def create(self, command: CreateAnexoCommand) -> RegisterAnexoResult:
        self._validate_actor(command.actor_id)
        self._enforce_size_limit(command)
        await self._load_expediente(command.expediente_id)
        created_at = datetime.now(UTC)
        anexo = self._build_anexo(command, created_at)
        try:
            with self._uow_factory():
                await self._anexo_repo.add(anexo)
                await self._audit_log.append(create_event(command, created_at))
                await self._audit_log.record_change(create_change(command, created_at))
        except RegisterAnexoError:
            raise
        except Exception as exc:
            raise RegisterAnexoError(
                f"anexo creation failed for {command.expediente_id!r}: {exc}"
            ) from exc
        return RegisterAnexoResult(command.anexo_id, command.expediente_id, created_at)

    @staticmethod
    def _build_anexo(command: CreateAnexoCommand, created_at: datetime) -> Anexo:
        anexo = Anexo(
            id=command.anexo_id,
            id_expediente=command.expediente_id,
            nombre=command.nombre,
            referencia=command.referencia,
            retention=command.retention,
            created_at=created_at,
        )
        anexo.record_event(AnexoEvent(kind="created", at=created_at, actor_id=command.actor_id))
        return anexo

    async def delete(self, command: DeleteAnexoCommand) -> DeleteAnexoResult:
        if command.actor_id is None:
            raise DeleteAnexoAuthorizationError("actor_id is required (deny-by-default)")
        anexo = await self._load_anexo(command.anexo_id)
        self._enforce_retention(anexo, context=f"delete:anexo-{anexo.id}")
        deleted_at = datetime.now(UTC)
        try:
            with self._uow_factory():
                await self._anexo_repo.delete(anexo.id)
                await self._audit_log.append(delete_event(command, anexo, deleted_at))
                await self._audit_log.record_change(delete_change(command, anexo, deleted_at))
        except DeleteAnexoError:
            raise
        except Exception as exc:
            raise DeleteAnexoError(f"anexo deletion failed for {anexo.id!r}: {exc}") from exc
        return DeleteAnexoResult(anexo.id, deleted_at)

    async def _load_expediente(self, expediente_id: Any) -> Expediente:
        loaded = await self._safe_get_expediente(expediente_id)
        if loaded is None:
            raise RegisterAnexoValidationError(f"expediente {expediente_id!r} not found")
        return cast(Expediente, loaded)

    async def _safe_get_expediente(self, expediente_id: Any) -> object | None:
        try:
            return await self._expediente_repo.get_by_id(expediente_id)
        except Exception as exc:
            raise RegisterAnexoError("expediente dependency unavailable") from exc

    async def _load_anexo(self, anexo_id: Any) -> Anexo:
        loaded = await self._safe_get_anexo(anexo_id)
        if loaded is None:
            raise RegisterAnexoValidationError(f"anexo {anexo_id!r} not found")
        return cast(Anexo, loaded)

    async def _safe_get_anexo(self, anexo_id: Any) -> object | None:
        try:
            return await self._anexo_repo.get_by_id(anexo_id)
        except Exception as exc:
            raise DeleteAnexoError("anexo dependency unavailable") from exc

    def _validate_actor(self, actor_id: Any) -> None:
        if actor_id is None:
            raise CreateAnexoAuthorizationError("actor_id is required (deny-by-default)")

    def _enforce_size_limit(self, command: CreateAnexoCommand) -> None:
        try:
            self._size_limit.ensure_within_limit(
                size_bytes=command.referencia.size_bytes,
                context=f"create:anexo-{command.anexo_id}",
            )
        except SizeLimitExceeded as exc:
            raise RegisterAnexoSizeLimitError(str(exc)) from exc

    def _enforce_retention(self, anexo: Anexo, *, context: str) -> None:
        created_at = anexo.created_at
        if created_at is None:
            return
        try:
            anexo.retention.enforce(
                at=datetime.now(UTC),
                created_at=created_at,
                context=context,
            )
        except RetentionExpired as exc:
            raise DeleteAnexoError(f"anexo retention expired: {exc}") from exc
