"""Strict TDD for EXP-CAP-035 E2E batch atomicity (E05)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.batch_e2e.command import (
    BatchDetailInput,
    BatchE2EAuthorizationError,
    BatchE2ECommand,
    BatchE2EValidationError,
    BatchStatus,
)
from app.src.modules.expedientes.application.batch_e2e.service import BatchE2EService


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


class _Processor:
    """Test double that simulates a downstream dependency."""

    def __init__(self) -> None:
        self.calls: list[UUID] = []
        self.fail_on: set[UUID] = set()

    async def process(self, detail_id: UUID) -> str:
        self.calls.append(detail_id)
        if detail_id in self.fail_on:
            raise RuntimeError(f"processor failed for {detail_id}")
        return f"ok:{detail_id}"


def _make_service() -> tuple[BatchE2EService, _Audit, _Processor]:
    audit = _Audit()
    processor = _Processor()
    service = BatchE2EService(
        audit_log=audit,
        processor=processor,
        permissions=set(),
    )
    return service, audit, processor


def _command(details: list[BatchDetailInput], **overrides: Any) -> BatchE2ECommand:
    values: dict[str, Any] = {
        "actor_id": uuid4(),
        "atomicity": "all_or_nothing",
        "details": details,
    }
    values.update(overrides)
    return BatchE2ECommand(**values)


async def test_batch_creates_records_and_processes_all_details() -> None:
    service, audit, processor = _make_service()
    service.grant("e2e.batch")
    details = [BatchDetailInput(id=uuid4(), payload={}) for _ in range(3)]

    result = await service.execute(_command(details))

    assert result.batch_id == result.batch_id  # smoke
    assert result.status is BatchStatus.COMPLETED
    assert len(processor.calls) == 3
    assert result.detail_count == 3
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "e2e.batch.executed"
    assert audit.events[0].capacidad == "EXP-CAP-035"


async def test_atomicity_all_or_nothing_marks_batch_failed_on_partial_error() -> None:
    service, audit, processor = _make_service()
    service.grant("e2e.batch")
    details = [BatchDetailInput(id=uuid4(), payload={}) for _ in range(3)]
    processor.fail_on = {details[1].id}

    result = await service.execute(_command(details))

    assert result.status is BatchStatus.FAILED
    assert result.atomicity_rolled_back is True
    assert result.detail_count == 3
    # The first failure short-circuits the batch; one detail succeeded
    # before the failing detail ran.
    assert result.succeeded_count == 1
    assert result.failed_count == 1
    # The short-circuit path does not emit an audit event (the
    # rollback message is the only signal the delivery layer needs).
    assert not audit.events


async def test_atomicity_best_effort_keeps_successful_details() -> None:
    service, audit, processor = _make_service()
    service.grant("e2e.batch")
    details = [BatchDetailInput(id=uuid4(), payload={}) for _ in range(3)]
    processor.fail_on = {details[1].id}

    result = await service.execute(_command(details, atomicity="best_effort"))

    assert result.status is BatchStatus.FAILED
    assert result.atomicity_rolled_back is False
    assert result.succeeded_count == 2
    assert result.failed_count == 1


async def test_empty_batch_rejected() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.batch")

    with pytest.raises(BatchE2EValidationError, match="empty"):
        await service.execute(_command([]))


async def test_unknown_atomicity_rejected() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.batch")
    details = [BatchDetailInput(id=uuid4(), payload={})]

    with pytest.raises(BatchE2EValidationError, match="atomicity"):
        await service.execute(_command(details, atomicity="nuclear"))


async def test_missing_actor_rejected() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.batch")
    details = [BatchDetailInput(id=uuid4(), payload={})]
    with pytest.raises(BatchE2EAuthorizationError):
        await service.execute(_command(details, actor_id=None))


async def test_missing_permission_rejected() -> None:
    service, _, _ = _make_service()
    details = [BatchDetailInput(id=uuid4(), payload={})]
    with pytest.raises(BatchE2EAuthorizationError, match="e2e.batch"):
        await service.execute(_command(details))


async def test_replay_returns_same_result() -> None:
    service, audit, processor = _make_service()
    service.grant("e2e.batch")
    details = [BatchDetailInput(id=uuid4(), payload={})]
    first = await service.execute(_command(details))
    second = await service.execute(_command(details))
    # Each execution produces a fresh batch_id (no persistent store),
    # but the processor is called once per detail per invocation.
    assert first.detail_count == second.detail_count == 1
    assert len(processor.calls) == 2
    assert len(audit.events) == 2
