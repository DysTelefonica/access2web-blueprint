"""ExpedienteEditService — CAP-002 edición concurrente (issue #227).

CAP-002 §Camino feliz: "se actualizan los campos manteniendo
invariantes y control de concurrencia, con salida determinista y
auditable". §Concurrencia o fallo: "preserva invariantes, no
duplica ni pierde evidencia, queda reintentable".

Implementation strategy:

1. **Validation** runs BEFORE the UoW opens — malformed input
   doesn't waste a transaction.
2. **Load** the current aggregate with the optimistic-lock version.
3. **Compute** the edit in memory (one or more field mutations).
4. **Validate invariants** via ``Expediente.__post_init__`` (F02).
5. **Open the UoW**. Inside:
   - UPDATE with ``WHERE version = ?`` (D-EXP-4) — the repository
     raises on stale version; we surface it as
     ``ExpedienteEditConflictError`` (409).
   - ``audit_log.append`` for the edit event.
   - One ``audit_log.record_change`` per changed field (C02 differs
     from C01: per-field changes, not one aggregate-wide change row).
6. **Commit** on success, **rollback** on any failure inside the
   ``with`` block.

Out of scope (documented in the PR body):

- Fields beyond ``estado`` and ``id_expediente_padre``: the
  ``Expediente`` aggregate (F02) doesn't carry ``titulo``/``importe``/
  ``fecha_inicio``/``fecha_fin``/``ordinal``. Adding them requires a
  follow-up that extends F02 first; C02 then commits per-field
  changes for those too.
- Read-model rebuild (CAP-001 §Camino feliz fourth item): tracked
  in #668 follow-up alongside vertical R.

DA-1: the application layer does NOT touch SQLAlchemy directly. It
goes through the UoW's session and the repositories' protocol
interface (DA-1).
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import replace
from datetime import UTC, datetime
from typing import Any, cast
from uuid import uuid4

from app.src.modules.expedientes.application.edit_expediente.command import (
    ExpedienteEditAuthorizationError,
    ExpedienteEditCommand,
    ExpedienteEditConflictError,
    ExpedienteEditError,
    ExpedienteEditResult,
    ExpedienteEditValidationError,
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


class ExpedienteEditService:
    """Use case: edit an expediente (CAP-002).

    Optimistic locking (D-EXP-4): the aggregate's ``version`` is
    supplied by the caller and matched against the stored version in
    the UPDATE. A mismatch raises ``ExpedienteEditConflictError``
    (HTTP 409) and rolls back the transaction.
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

    async def execute(self, command: ExpedienteEditCommand) -> ExpedienteEditResult:
        """Run the edit end-to-end.

        Raises:
            ExpedienteEditValidationError: malformed input or unknown aggregate.
            ExpedienteEditAuthorizationError: actor missing (D-EXP-3).
            ExpedienteEditConflictError: version stale (D-EXP-4).
            ExpedienteEditError: any other persistence failure.
        """
        self._validate(command)

        loaded = await self._expediente_repo.get_by_id(command.expediente_id)
        if loaded is None:
            raise ExpedienteEditValidationError(f"expediente {command.expediente_id!r} not found")
        # The protocol returns ``object`` for portability; this use
        # case knows the concrete type.
        existing: Expediente = cast(Expediente, loaded)

        # Version check FIRST, before computing the diff. A stale
        # version is a conflict regardless of whether the supplied
        # fields would actually change anything (D-EXP-4).
        if getattr(existing, "version", None) != command.expected_version:
            raise ExpedienteEditConflictError(
                f"version stale: stored={getattr(existing, 'version', None)}, "
                f"incoming={command.expected_version}"
            )

        # Compute the edited aggregate + the list of changed fields.
        edited, fields_changed = self._apply_edit(existing, command)

        # No-op edit: the supplied fields equal the stored values.
        # Nothing to do — no audit, no change rows, no commit. This
        # matches the conventional "if nothing changed, nothing to
        # record" pattern (audit logs grow on actual changes only).
        if not fields_changed:
            return ExpedienteEditResult(
                expediente_id=existing.id,
                new_version=existing.version,
                modified_at=datetime.now(UTC),
                fields_changed=(),
            )

        # ``dataclasses.replace`` invokes ``__post_init__``, so the
        # aggregate invariants were already checked by ``_apply_edit``.
        modified_at = datetime.now(UTC)

        try:
            uow = self._uow_factory()
            with uow as _session:
                # The version check moves inside the UoW so a stale
                # version rolls back the transaction (D-EXP-4 audit-
                # in-same-tx). The repo enforces the conditional UPDATE
                # (``WHERE version = expected_version``); if 0 rows
                # affected, raise ``ExpedienteEditConflictError``.
                if getattr(existing, "version", None) != command.expected_version:
                    raise ExpedienteEditConflictError(
                        f"version stale: stored={getattr(existing, 'version', None)}, "
                        f"incoming={command.expected_version}"
                    )
                await self._expediente_repo.update(edited)
                await self._audit_log.append(
                    self._audit_event(
                        command=command,
                        modified_at=modified_at,
                    )
                )
                for field_name, old_value, new_value in fields_changed:
                    await self._audit_log.record_change(
                        self._change_record(
                            command=command,
                            field_name=field_name,
                            old_value=old_value,
                            new_value=new_value,
                            modified_at=modified_at,
                        )
                    )
        except ExpedienteEditError:
            raise
        except Exception as exc:
            raise ExpedienteEditError(
                f"edit failed for expediente {command.expediente_id!r}: {exc}"
            ) from exc

        return ExpedienteEditResult(
            expediente_id=edited.id,
            new_version=edited.version,
            modified_at=modified_at,
            fields_changed=tuple(name for name, _, _ in fields_changed),
        )

    @staticmethod
    def _validate(command: ExpedienteEditCommand) -> None:
        """Check invariants before opening the UoW.

        A missing actor is a deny-by-default violation (D-EXP-3).
        A missing expected_version is a malformed input (the caller
        didn't supply the optimistic-lock token).
        """
        if command.expediente_id is None:
            raise ExpedienteEditValidationError("expediente_id is required")
        if command.expected_version is None or command.expected_version < 1:
            raise ExpedienteEditValidationError(
                "expected_version is required and must be >= 1 (D-EXP-4)"
            )
        if command.actor_id is None:
            raise ExpedienteEditAuthorizationError("actor_id is required (D-EXP-3 deny-by-default)")

    @staticmethod
    def _apply_edit(
        existing: Expediente,
        command: ExpedienteEditCommand,
    ) -> tuple[Expediente, list[tuple[str, Any, Any]]]:
        """Apply the field-level mutations and bump the version.

        Returns the new aggregate and a list of (field, old, new)
        tuples for the change log. Bumps the version by 1 on any
        successful edit (D-EXP-4).
        """
        fields_changed: list[tuple[str, Any, Any]] = []

        if command.estado is not None and command.estado != existing.estado:
            fields_changed.append(("estado", existing.estado.value, command.estado.value))
            new_estado = command.estado
        else:
            new_estado = existing.estado

        if (
            command.id_expediente_padre is not None
            and command.id_expediente_padre != existing.id_expediente_padre
        ):
            fields_changed.append(
                (
                    "id_expediente_padre",
                    existing.id_expediente_padre,
                    command.id_expediente_padre,
                )
            )
            new_padre = command.id_expediente_padre
        else:
            new_padre = existing.id_expediente_padre

        if not fields_changed:
            # No-op edit: caller supplied fields that equal the stored
            # values. The ``execute`` wrapper detects this and returns
            # silently — see the early-return above.
            return existing, []

        # Bump version (D-EXP-4 optimistic locking).
        new_version = existing.version + 1

        edited = replace(
            existing,
            estado=new_estado,
            id_expediente_padre=new_padre,
            version=new_version,
            updated_at=datetime.now(UTC),
        )
        return edited, fields_changed

    @staticmethod
    def _audit_event(
        *,
        command: ExpedienteEditCommand,
        modified_at: datetime,
    ) -> ExpedienteAuditEvent:
        """Build the audit event for the edit.

        Mirrors ``ExpedienteAuditEvent`` (F01) without coupling this
        module to that import at module level — the test fakes can
        inject their own ``AuditLogPort``.
        """
        return ExpedienteAuditEvent(
            id=uuid4(),
            event_type="expediente.updated",
            actor_id=command.actor_id,
            target_id=command.expediente_id,
            capacidad="EXP-CAP-002",
            module="expedientes",
            result="ok",
        )

    @staticmethod
    def _change_record(
        *,
        command: ExpedienteEditCommand,
        field_name: str,
        old_value: Any,
        new_value: Any,
        modified_at: datetime,
    ) -> ChangeRecord:
        """Build one ``ChangeRecord`` per modified field.

        CAP-002 emits one row per field change (not one aggregate-
        wide row like CAP-001). The ``nombre_tabla`` matches
        ``TbExpedientes``; ``valor_inicial``/``valor_final`` are string
        representations of the typed values.
        """
        return ChangeRecord(
            id=uuid4(),
            nombre_tabla="expedientes",
            id_expediente=command.expediente_id,
            nombre_campo=field_name,
            valor_inicial=None if old_value is None else str(old_value),
            valor_final=None if new_value is None else str(new_value),
            fecha_cambio=modified_at,
            id_usuario_cambio=command.actor_id,
            accion="edit",
        )
