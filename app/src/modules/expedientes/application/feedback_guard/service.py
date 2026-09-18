"""EXP-CAP-032 idempotent feedback guard service."""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, cast
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.feedback_guard.command import (
    FeedbackGuardAuthorizationError,
    FeedbackGuardCommand,
    FeedbackGuardError,
    FeedbackGuardState,
    FeedbackGuardStatus,
    FeedbackGuardValidationError,
)
from app.src.modules.expedientes.ports.audit_log import AuditLogPort, ExpedienteAuditEvent
from app.src.modules.expedientes.ports.idempotency import IdempotencyPort, IdempotencyRecord


class FeedbackGuardService:
    """Open the BUSY state and persist replay-safe idempotency for an
    incoming mutation. The originating use case (autosave, edit, etc.)
    hands the resulting ``FeedbackGuardState`` to the delivery layer
    and updates it on completion by calling ``mark_finished`` or
    ``mark_error``."""

    def __init__(
        self,
        *,
        idempotency: IdempotencyPort,
        uow_factory: Callable[[], Any],
        audit_log: AuditLogPort | None = None,
    ) -> None:
        self._idempotency = idempotency
        self._uow_factory = uow_factory
        self._audit_log = audit_log

    async def execute(self, command: FeedbackGuardCommand) -> FeedbackGuardState:
        self._validate(command)
        try:
            recorded = await self._idempotency.get(command.idempotency_key)
        except Exception as exc:
            raise FeedbackGuardError(
                f"feedback guard failed to read idempotency for {command.idempotency_key!r}: {exc}"
            ) from exc

        if recorded is not None:
            state = cast(FeedbackGuardState, recorded.result)
            if state.status is FeedbackGuardStatus.BUSY:
                state = FeedbackGuardState(
                    idempotency_key=state.idempotency_key,
                    operation=state.operation,
                    status=FeedbackGuardStatus.FINISHED,
                    completed_at=datetime.now(UTC),
                    result=state.result,
                )
            return state

        state = FeedbackGuardState(
            idempotency_key=command.idempotency_key,
            operation=command.operation,
            status=FeedbackGuardStatus.BUSY,
            completed_at=None,
            result=None,
        )
        try:
            with self._uow_factory():
                if self._audit_log is not None:
                    await self._audit_log.append(self._busy_event(command))
                await self._idempotency.record(
                    IdempotencyRecord(
                        command.idempotency_key,
                        f"{command.operation}|{command.idempotency_key}",
                        state,
                    )
                )
        except FeedbackGuardError:
            raise
        except Exception as exc:
            raise FeedbackGuardError(
                f"feedback guard failed for {command.idempotency_key!r}: {exc}"
            ) from exc
        return state

    async def mark_finished(
        self,
        *,
        idempotency_key: UUID,
        operation: str,
        result: Any,
    ) -> FeedbackGuardState:
        state = FeedbackGuardState(
            idempotency_key=idempotency_key,
            operation=operation,
            status=FeedbackGuardStatus.FINISHED,
            completed_at=datetime.now(UTC),
            result=result,
        )
        await self._idempotency.record(IdempotencyRecord(idempotency_key, operation, state))
        return state

    async def mark_error(
        self,
        *,
        idempotency_key: UUID,
        operation: str,
        detail: str,
    ) -> FeedbackGuardState:
        state = FeedbackGuardState(
            idempotency_key=idempotency_key,
            operation=operation,
            status=FeedbackGuardStatus.ERROR,
            completed_at=datetime.now(UTC),
            result=detail,
        )
        await self._idempotency.record(IdempotencyRecord(idempotency_key, operation, state))
        return state

    @staticmethod
    def _validate(command: FeedbackGuardCommand) -> None:
        if command.actor_id is None:
            raise FeedbackGuardAuthorizationError("actor_id is required (deny-by-default)")
        if not command.operation or not command.operation.strip():
            raise FeedbackGuardValidationError("operation label is required")
        if command.idempotency_key is None:
            raise FeedbackGuardValidationError("idempotency_key is required")

    @staticmethod
    def _busy_event(command: FeedbackGuardCommand) -> ExpedienteAuditEvent:
        return ExpedienteAuditEvent(
            id=uuid4(),
            event_type="feedback.busy",
            actor_id=command.actor_id,
            target_id=command.idempotency_key,
            capacidad="EXP-CAP-032",
            payload={"operation": command.operation},
        )
