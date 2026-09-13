"""Transactional EXP-CAP-014 suministrador registration."""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast

from app.src.modules.expedientes.application.register_suministrador._evidence import (
    audit_event,
    change_record,
)
from app.src.modules.expedientes.application.register_suministrador.command import (
    RegisterSuministradorAuthorizationError,
    RegisterSuministradorCommand,
    RegisterSuministradorConflictError,
    RegisterSuministradorError,
    RegisterSuministradorResult,
    RegisterSuministradorValidationError,
)
from app.src.modules.expedientes.domain.suministrador import Suministrador
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort
from app.src.modules.expedientes.ports.suministrador_repository import SuministradorRepositoryPort


class RegisterSuministradorService:
    """Register one suministrador and its evidence in the same UoW."""

    def __init__(
        self,
        *,
        expediente_repo: ExpedienteRepositoryPort,
        suministrador_repo: SuministradorRepositoryPort,
        audit_log: AuditLogPort,
        uow_factory: Callable[[], Any],
    ) -> None:
        self._expediente_repo = expediente_repo
        self._suministrador_repo = suministrador_repo
        self._audit_log = audit_log
        self._uow_factory = uow_factory

    async def execute(self, command: RegisterSuministradorCommand) -> RegisterSuministradorResult:
        self._validate(command)
        existing = await self._load_existing(command)
        if existing is not None:
            if not self._same_payload(existing, command):
                raise RegisterSuministradorConflictError(
                    f"suministrador {command.suministrador_id!r} "
                    "already exists with different payload"
                )
            return RegisterSuministradorResult(
                existing.id,
                existing.id_expediente,
                existing.created_at or datetime.now(UTC),
            )

        try:
            loaded = await self._expediente_repo.get_by_id(command.expediente_id)
        except Exception as exc:
            raise RegisterSuministradorError("expediente dependency unavailable") from exc
        if loaded is None:
            raise RegisterSuministradorValidationError(
                f"expediente {command.expediente_id!r} not found"
            )
        self._reject_self_loop(command)
        await self._enforce_tree_invariant(command)
        registered_at = datetime.now(UTC)
        try:
            suministrador = Suministrador(
                id=command.suministrador_id,
                id_expediente=command.expediente_id,
                id_suministrador=command.id_suministrador,
                id_padre=command.id_padre,
                descripcion=command.descripcion,
                contratista_principal=command.contratista_principal,
                sub_contratista=command.sub_contratista,
                created_at=registered_at,
            )
        except ValueError as exc:
            raise RegisterSuministradorValidationError(str(exc)) from exc

        try:
            with self._uow_factory():
                await self._suministrador_repo.upsert(suministrador)
                await self._audit_log.append(audit_event(command, registered_at))
                await self._audit_log.record_change(change_record(command, registered_at))
        except RegisterSuministradorError:
            raise
        except Exception as exc:
            raise RegisterSuministradorError(
                f"suministrador registration failed for {command.expediente_id!r}: {exc}"
            ) from exc

        return RegisterSuministradorResult(
            command.suministrador_id, command.expediente_id, registered_at
        )

    async def _enforce_tree_invariant(self, command: RegisterSuministradorCommand) -> None:
        """Apply CAP-014 «el árbol manda».

        - A root row (no parent) must mark ``contratista_principal``.
        - A child row (parent set) must mark ``sub_contratista``.
        - The parent must exist in the same expediente and must not be
          the row itself.
        """
        if command.id_padre is None:
            if command.contratista_principal is not True:
                raise RegisterSuministradorValidationError(
                    "root suministrador must set contratista_principal"
                )
            return
        if command.sub_contratista is not True:
            raise RegisterSuministradorValidationError(
                "child suministrador must set sub_contratista and provide id_padre"
            )
        try:
            parent = await self._suministrador_repo.get_by_id(command.id_padre)
        except Exception as exc:
            raise RegisterSuministradorError("suministrador dependency unavailable") from exc
        if parent is None:
            raise RegisterSuministradorValidationError(
                f"parent suministrador {command.id_padre!r} not found"
            )
        if getattr(parent, "id_expediente", None) != command.expediente_id:
            raise RegisterSuministradorValidationError("parent must belong to the same expediente")

    async def _load_existing(self, command: RegisterSuministradorCommand) -> Suministrador | None:
        try:
            return cast(
                Suministrador | None,
                await self._suministrador_repo.get_by_id(command.suministrador_id),
            )
        except Exception as exc:
            raise RegisterSuministradorError("suministrador dependency unavailable") from exc

    @staticmethod
    def _validate(command: RegisterSuministradorCommand) -> None:
        if command.actor_id is None:
            raise RegisterSuministradorAuthorizationError("actor_id is required (deny-by-default)")
        if (
            command.suministrador_id is None
            or command.expediente_id is None
            or command.id_suministrador is None
        ):
            raise RegisterSuministradorValidationError(
                "suministrador_id, expediente_id and id_suministrador are required"
            )

    @staticmethod
    def _reject_self_loop(command: RegisterSuministradorCommand) -> None:
        """A row whose parent equals its own row id would form a 1-cycle.

        ``id_suministrador`` is the catalog FK; the row id is the
        ``suministrador_id`` PK. If the caller reuses the catalog id as
        the row id, the cycle check fires. The parent supplied as the
        existing root row id matches ``suministrador_id``, so the
        check rejects the cycle before reaching the repository.
        """
        if command.id_padre is not None and command.id_padre == command.suministrador_id:
            raise RegisterSuministradorValidationError(
                "id_padre cannot equal the row id (cycle detected)"
            )

    @staticmethod
    def _same_payload(existing: Suministrador, command: RegisterSuministradorCommand) -> bool:
        return (
            existing.id_expediente,
            existing.id_suministrador,
            existing.id_padre,
            existing.descripcion,
            existing.contratista_principal,
            existing.sub_contratista,
        ) == (
            command.expediente_id,
            command.id_suministrador,
            command.id_padre,
            command.descripcion,
            command.contratista_principal,
            command.sub_contratista,
        )
