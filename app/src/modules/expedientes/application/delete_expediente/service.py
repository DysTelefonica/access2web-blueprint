"""ExpedienteDeleteService — CAP-003 baja condicionada (issue #228).

CAP-003 §Camino feliz: "se elimina el agregado de forma auditable
sin afectar relaciones no incluidas". §Concurrencia o fallo:
"preserva invariantes, no duplica ni pierde evidencia, queda
reintentable".

Implementation strategy:

1. **Validation** runs BEFORE the UoW opens — malformed input
   doesn't waste a transaction.
2. **Load** the aggregate; **validate version** (D-EXP-4) pre-UoW.
3. **Check children** unless ``force=True`` — CAP-003 "impedir
   pérdida de hijos".
4. **Open the UoW**. Inside:
   - DELETE the aggregate (the repo is responsible for the
     conditional DELETE — the service has already validated version
     and the children-presence decision).
   - ``audit_log.append`` for the deletion event.
   - ``audit_log.record_change`` with ``accion="delete"`` carrying the
     pre-delete version.
5. **Commit** on success, **rollback** on any failure inside the
   ``with`` block.

DA-1: the application layer does NOT touch SQLAlchemy directly. It
goes through the UoW's session and the repositories' protocol
interface.
"""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast
from uuid import uuid4

from app.src.modules.expedientes.application.delete_expediente.command import (
    ExpedienteDeleteAuthorizationError,
    ExpedienteDeleteCommand,
    ExpedienteDeleteConflictError,
    ExpedienteDeleteError,
    ExpedienteDeleteResult,
    ExpedienteDeleteValidationError,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.ports.audit_log import (
    AuditLogPort,
    ChangeRecord,
    ExpedienteAuditEvent,
)
from app.src.modules.expedientes.ports.expediente_repository import (
    ExpedienteRepositoryPort,
)
from app.src.modules.expedientes.ports.hito_repository import HitoRepositoryPort


class ExpedienteDeleteService:
    """Use case: delete an expediente (CAP-003).

    The conditional aspect: by default the use case refuses the
    delete when the aggregate has children. The caller can override
    by passing ``force=True`` to the command.
    """

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

    async def execute(self, command: ExpedienteDeleteCommand) -> ExpedienteDeleteResult:
        """Run the delete end-to-end.

        Raises:
            ExpedienteDeleteValidationError: malformed input, unknown
                aggregate, or children-presence refused.
            ExpedienteDeleteAuthorizationError: actor missing.
            ExpedienteDeleteConflictError: version stale.
            ExpedienteDeleteError: any other persistence failure.
        """
        self._validate(command)

        loaded = await self._expediente_repo.get_by_id(command.expediente_id)
        if loaded is None:
            raise ExpedienteDeleteValidationError(f"expediente {command.expediente_id!r} not found")
        existing: Expediente = cast(Expediente, loaded)

        # Version check first (D-EXP-4). A stale version is a
        # conflict regardless of whether the aggregate has children.
        if getattr(existing, "version", None) != command.expected_version:
            raise ExpedienteDeleteConflictError(
                f"version stale: stored={getattr(existing, 'version', None)}, "
                f"incoming={command.expected_version}"
            )

        # Conditional check: refuse if there are children (CAP-003
        # "impedir pérdida de hijos"), unless ``force=True``.
        if not command.force:
            has_children = await self._expediente_repo.has_children(command.expediente_id)
            if has_children:
                raise ExpedienteDeleteValidationError(
                    f"expediente {command.expediente_id!r} has children; "
                    "pass force=True to override"
                )

        deleted_at = datetime.now(UTC)

        try:
            uow = self._uow_factory()
            with uow as _session:
                # The repo enforces the conditional DELETE in
                # production (``DELETE ... WHERE id = ? AND version = ?``).
                await self._expediente_repo.delete(command.expediente_id)
                await self._audit_log.append(
                    self._audit_event(
                        command=command,
                        deleted_at=deleted_at,
                    )
                )
                await self._audit_log.record_change(
                    self._change_record(
                        command=command,
                        existing=existing,
                        deleted_at=deleted_at,
                    )
                )
        except ExpedienteDeleteError:
            raise
        except Exception as exc:
            raise ExpedienteDeleteError(
                f"delete failed for expediente {command.expediente_id!r}: {exc}"
            ) from exc

        return ExpedienteDeleteResult(
            expediente_id=command.expediente_id,
            deleted_at=deleted_at,
        )

    @staticmethod
    def _validate(command: ExpedienteDeleteCommand) -> None:
        if command.expediente_id is None:
            raise ExpedienteDeleteValidationError("expediente_id is required")
        if command.expected_version is None or command.expected_version < 1:
            raise ExpedienteDeleteValidationError(
                "expected_version is required and must be >= 1 (D-EXP-4)"
            )
        if command.actor_id is None:
            raise ExpedienteDeleteAuthorizationError(
                "actor_id is required (D-EXP-3 deny-by-default)"
            )

    @staticmethod
    def _audit_event(
        *,
        command: ExpedienteDeleteCommand,
        deleted_at: datetime,
    ) -> ExpedienteAuditEvent:
        return ExpedienteAuditEvent(
            id=uuid4(),
            event_type="expediente.deleted",
            actor_id=command.actor_id,
            target_id=command.expediente_id,
            capacidad="EXP-CAP-003",
            module="expedientes",
            result="ok",
        )

    @staticmethod
    def _change_record(
        *,
        command: ExpedienteDeleteCommand,
        existing: Expediente,
        deleted_at: datetime,
    ) -> ChangeRecord:
        """One ``ChangeRecord`` row for the deletion.

        ``valor_inicial`` carries the pre-delete version (useful for
        audit). ``valor_final`` is empty — the row is gone. The
        ``accion`` is ``delete`` (lower-case) to match C01/C02's
        convention.
        """
        return ChangeRecord(
            id=uuid4(),
            nombre_tabla="expedientes",
            id_expediente=command.expediente_id,
            nombre_campo=None,
            valor_inicial=str(getattr(existing, "version", None)),
            valor_final=None,
            fecha_cambio=deleted_at,
            id_usuario_cambio=command.actor_id,
            accion="delete",
        )
