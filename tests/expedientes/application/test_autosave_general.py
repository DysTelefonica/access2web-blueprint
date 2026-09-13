"""Strict TDD for CAP-030 autosave core (#234)."""

from dataclasses import replace
from datetime import date
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.autosave_general.command import (
    AutosaveGeneralCommand,
    AutosaveGeneralConflictError,
    AutosaveGeneralError,
    AutosaveGeneralValidationError,
)
from app.src.modules.expedientes.application.autosave_general.service import AutosaveGeneralService
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo
from app.src.modules.expedientes.ports.idempotency import IdempotencyRecord


class _State:
    def __init__(self, aggregate: Expediente) -> None:
        self.aggregate, self.updates, self.fail_update = aggregate, 0, False
        self.records: dict[UUID, IdempotencyRecord] = {}
        self.events: list[Any] = []
        self.changes: list[Any] = []

    async def get_by_id(self, expediente_id: UUID) -> object | None:
        return self.aggregate if expediente_id == self.aggregate.id else None

    async def update(self, aggregate: Expediente) -> object:
        self.updates += 1
        if self.fail_update:
            raise RuntimeError("write failed")
        self.aggregate = aggregate
        return aggregate

    async def get(self, key: UUID) -> IdempotencyRecord | None:
        return self.records.get(key)

    async def record(self, record: IdempotencyRecord) -> None:
        self.records[record.key] = record

    async def append(self, event: Any) -> None:
        self.events.append(event)

    async def record_change(self, change: Any) -> None:
        self.changes.append(change)


class _UoW:
    def __init__(self) -> None:
        self.entries = self.commits = self.rollbacks = 0

    def __enter__(self) -> object:
        self.entries += 1
        return object()

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        self.commits += exc_type is None
        self.rollbacks += exc_type is not None


def _fixture() -> tuple[AutosaveGeneralService, _State, _UoW, Expediente]:
    aggregate = Expediente(
        id=uuid4(), tipo=ExpedienteTipo.AM, estado=ExpedienteEstado.BORRADOR, version=1
    )
    state, uow = _State(aggregate), _UoW()
    service = AutosaveGeneralService(
        expediente_repo=state,
        idempotency=state,
        audit_log=state,
        uow_factory=lambda: uow,
    )
    return service, state, uow, aggregate


def _command(aggregate: Expediente) -> AutosaveGeneralCommand:
    return AutosaveGeneralCommand(
        expediente_id=aggregate.id,
        expected_version=1,
        idempotency_key=uuid4(),
        fecha_inicio_contrato=date(2026, 1, 1),
        fecha_fin_contrato=date(2026, 12, 31),
        actor_id=uuid4(),
    )


async def test_autosave_persists_dates_audit_and_key_in_one_uow() -> None:
    service, state, uow, aggregate = _fixture()
    command = _command(aggregate)
    result = await service.execute(command)
    assert result.new_version == 2
    assert state.aggregate.fecha_inicio_contrato == date(2026, 1, 1)
    assert state.aggregate.fecha_fin_contrato == date(2026, 12, 31)
    assert uow.commits == 1
    assert len(state.events) == 1 and len(state.changes) == 2
    assert state.records[command.idempotency_key].result == result


async def test_exact_retry_returns_recorded_result_without_new_write() -> None:
    service, state, uow, aggregate = _fixture()
    command = _command(aggregate)
    first = await service.execute(command)
    assert await service.execute(command) == first
    assert state.updates == 1 and uow.entries == 1 and len(state.events) == 1


async def test_reused_key_with_different_payload_is_conflict() -> None:
    service, _, _, aggregate = _fixture()
    command = _command(aggregate)
    await service.execute(command)
    with pytest.raises(AutosaveGeneralConflictError, match="different payload"):
        await service.execute(replace(command, fecha_fin_contrato=date(2027, 1, 1)))


async def test_stale_version_has_no_transactional_effect() -> None:
    service, state, uow, aggregate = _fixture()
    with pytest.raises(AutosaveGeneralConflictError, match="version stale"):
        await service.execute(replace(_command(aggregate), expected_version=9))
    assert state.updates == 0 and not state.records and not state.events and uow.entries == 0


async def test_failed_write_rolls_back_without_recording_the_key() -> None:
    service, state, uow, aggregate = _fixture()
    state.fail_update = True
    with pytest.raises(AutosaveGeneralError, match="autosave failed"):
        await service.execute(_command(aggregate))
    assert uow.rollbacks == 1 and not state.records and not state.events


async def test_invalid_date_range_is_rejected_before_uow() -> None:
    service, state, uow, aggregate = _fixture()
    command = replace(
        _command(aggregate),
        fecha_inicio_contrato=date(2026, 12, 31),
        fecha_fin_contrato=date(2026, 1, 1),
    )
    with pytest.raises(AutosaveGeneralValidationError, match="before"):
        await service.execute(command)
    assert state.updates == 0 and not state.records and uow.entries == 0
