"""Service implementation for EXP-CAP-035 E2E batch atomicity (E05)."""

from __future__ import annotations

from typing import Any
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.batch_e2e._evidence import (
    batch_event,
)
from app.src.modules.expedientes.application.batch_e2e.command import (
    BatchAtomicity,
    BatchE2EAuthorizationError,
    BatchE2ECommand,
    BatchE2EResult,
    BatchE2EValidationError,
    BatchStatus,
)


class BatchE2EService:
    """Application-level E2E batch execution (CAP-035).

    The service runs every detail through the injected ``processor``.
    For ``all_or_nothing`` atomicity the first failure short-circuits
    the batch and marks it FAILED. For ``best_effort`` atomicity the
    batch keeps the partial result and reports the failure count.
    """

    def __init__(
        self,
        *,
        audit_log: Any,
        processor: Any,
        permissions: set[str],
    ) -> None:
        self._audit_log = audit_log
        self._processor = processor
        self._permissions = permissions

    def grant(self, permission: str) -> None:
        self._permissions.add(permission)

    async def execute(self, command: BatchE2ECommand) -> BatchE2EResult:
        self._check_actor(command.actor_id)
        if not command.details:
            raise BatchE2EValidationError("batch is empty")
        if command.atomicity not in {a.value for a in BatchAtomicity}:
            raise BatchE2EValidationError(
                f"unknown atomicity {command.atomicity!r}; "
                f"expected one of {[a.value for a in BatchAtomicity]!r}"
            )

        atomicity = BatchAtomicity(command.atomicity)
        succeeded: list[UUID] = []
        failed: list[UUID] = []
        for detail in command.details:
            try:
                await self._processor.process(detail.id)
            except Exception:
                failed.append(detail.id)
                if atomicity is BatchAtomicity.ALL_OR_NOTHING:
                    return BatchE2EResult(
                        batch_id=uuid4(),
                        status=BatchStatus.FAILED,
                        detail_count=len(command.details),
                        succeeded_count=len(succeeded),
                        failed_count=len(failed),
                        atomicity_rolled_back=True,
                    )
            else:
                succeeded.append(detail.id)

        status = BatchStatus.COMPLETED if not failed else BatchStatus.FAILED
        result = BatchE2EResult(
            batch_id=uuid4(),
            status=status,
            detail_count=len(command.details),
            succeeded_count=len(succeeded),
            failed_count=len(failed),
            atomicity_rolled_back=False,
        )
        assert command.actor_id is not None
        await self._audit_log.append(batch_event(command.actor_id, result))
        return result

    def _check_actor(self, actor_id: UUID | None) -> None:
        if actor_id is None:
            raise BatchE2EAuthorizationError("actor_id is required (deny-by-default)")
        if "e2e.batch" not in self._permissions:
            raise BatchE2EAuthorizationError("actor lacks permission 'e2e.batch'")


__all__ = ["BatchE2EService"]
