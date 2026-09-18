"""Strict TDD for EXP-CAP-032 idempotent feedback / double-submit guard."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.feedback_guard.command import (
    FeedbackGuardAuthorizationError,
    FeedbackGuardCommand,
    FeedbackGuardError,
    FeedbackGuardStatus,
    FeedbackGuardValidationError,
)
from app.src.modules.expedientes.application.feedback_guard.service import (
    FeedbackGuardService,
)


class _Idempotency:
    def __init__(self) -> None:
        self.records: dict[UUID, Any] = {}

    async def get(self, key: UUID) -> object | None:
        return self.records.get(key)

    async def record(self, record: Any) -> None:
        self.records[record.key] = record


@dataclass
class _UoW:
    commits: int = 0
    rollbacks: int = 0

    def __enter__(self) -> object:
        return object()

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        if exc_type is None:
            self.commits += 1
        else:
            self.rollbacks += 1


def _make_service() -> tuple[FeedbackGuardService, _Idempotency, _UoW]:
    idemp = _Idempotency()
    uow = _UoW()
    service = FeedbackGuardService(
        idempotency=idemp,
        uow_factory=lambda: uow,
    )
    return service, idemp, uow


def _command(**overrides: Any) -> FeedbackGuardCommand:
    values: dict[str, Any] = {
        "idempotency_key": uuid4(),
        "operation": "autosave_general",
        "actor_id": uuid4(),
    }
    values.update(overrides)
    return FeedbackGuardCommand(**values)


async def test_open_emits_busy_state_audit() -> None:
    service, idemp, uow = _make_service()
    command = _command()

    result = await service.execute(command)

    assert result.status is FeedbackGuardStatus.BUSY
    assert result.operation == "autosave_general"
    assert command.idempotency_key in idemp.records
    assert uow.commits == 1


async def test_open_records_progress_event_in_audit_port() -> None:
    events: list[Any] = []

    class _Audit:
        async def append(self, event: Any) -> None:
            events.append(event)

    service = FeedbackGuardService(
        idempotency=_Idempotency(),
        uow_factory=lambda: _UoW(),
        audit_log=_Audit(),
    )
    await service.execute(_command())

    assert len(events) == 1
    assert events[0].event_type == "feedback.busy"


async def test_replay_emits_finished_state_with_recorded_result() -> None:
    service, idemp, _ = _make_service()
    command = _command()

    first = await service.execute(command)
    second = await service.execute(command)

    assert second.status is FeedbackGuardStatus.FINISHED
    assert second.completed_at is not None
    assert second.result == first.result


async def test_double_submit_without_audit_does_not_flip_state() -> None:
    """A racing double-submit before the audit event lands still resolves
    deterministically: the second submit observes the BUSY state and the
    service refuses to transition until the first one finishes.
    """
    events: list[Any] = []
    idemp = _Idempotency()

    class _Audit:
        async def append(self, event: Any) -> None:
            events.append(event)

    service = FeedbackGuardService(
        idempotency=idemp,
        uow_factory=lambda: _UoW(),
        audit_log=_Audit(),
    )

    key = uuid4()
    cmd_a = _command(idempotency_key=key)
    cmd_b = _command(idempotency_key=key)

    first = await service.execute(cmd_a)
    second = await service.execute(cmd_b)

    assert first.status is FeedbackGuardStatus.BUSY
    assert second.status is FeedbackGuardStatus.FINISHED
    assert second.completed_at is not None
    assert len(events) == 1  # BUSY recorded once; replay returns FINISHED


async def test_denies_missing_actor() -> None:
    service, *_ = _make_service()

    with pytest.raises(FeedbackGuardAuthorizationError, match="actor_id"):
        await service.execute(_command(actor_id=None))


async def test_rejects_missing_operation_label() -> None:
    service, *_ = _make_service()

    with pytest.raises(FeedbackGuardValidationError, match="operation"):
        await service.execute(_command(operation=""))


async def test_dependency_failure_marks_error_state_and_rolls_back() -> None:
    class _BrokenIdempotency:
        async def get(self, key: UUID) -> object | None:
            raise RuntimeError("database unavailable")

        async def record(self, record: Any) -> None:
            raise RuntimeError("database unavailable")

    service = FeedbackGuardService(
        idempotency=_BrokenIdempotency(),
        uow_factory=lambda: _UoW(),
    )

    with pytest.raises(FeedbackGuardError, match="feedback guard failed"):
        await service.execute(_command())
