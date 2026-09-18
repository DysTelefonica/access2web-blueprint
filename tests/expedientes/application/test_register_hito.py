"""Strict TDD for EXP-CAP-008 milestone registration."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.register_hito.command import (
    RegisterHitoAuthorizationError,
    RegisterHitoCommand,
    RegisterHitoConflictError,
    RegisterHitoError,
    RegisterHitoValidationError,
)
from app.src.modules.expedientes.application.register_hito.service import RegisterHitoService
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


class _ExpedienteRepo:
    def __init__(self, expediente: Expediente) -> None:
        self.expediente = expediente

    async def get_by_id(self, expediente_id: UUID) -> object | None:
        return self.expediente if expediente_id == self.expediente.id else None


class _HitoRepo:
    def __init__(self) -> None:
        self.by_id: dict[UUID, Any] = {}
        self.fail = False

    async def get_by_id(self, hito_id: UUID) -> object | None:
        return self.by_id.get(hito_id)

    async def upsert(self, hito: Any) -> object:
        if self.fail:
            raise RuntimeError("database unavailable")
        self.by_id[hito.id] = hito
        return hito


class _AuditLog:
    def __init__(self) -> None:
        self.events: list[object] = []
        self.changes: list[object] = []

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


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


def _make_service() -> tuple[RegisterHitoService, _HitoRepo, _AuditLog, _UoW, UUID]:
    expediente = Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
    )
    hito_repo = _HitoRepo()
    audit = _AuditLog()
    uow = _UoW()
    service = RegisterHitoService(
        expediente_repo=_ExpedienteRepo(expediente),
        hito_repo=hito_repo,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    return service, hito_repo, audit, uow, expediente.id


def _command(expediente_id: UUID, **overrides: Any) -> RegisterHitoCommand:
    values: dict[str, Any] = {
        "hito_id": uuid4(),
        "expediente_id": expediente_id,
        "descripcion": "Formalización",
        "fecha_hito": date(2026, 10, 1),
        "garantia_fecha_fin": date(2027, 10, 1),
        "importe": Decimal("1250.50"),
        "actor_id": uuid4(),
    }
    values.update(overrides)
    return RegisterHitoCommand(**values)


async def test_registers_auditable_hito_atomically() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)

    result = await service.execute(command)

    stored = repo.by_id[command.hito_id]
    assert result.hito_id == command.hito_id
    assert stored.id_expediente == expediente_id
    assert stored.descripcion == "Formalización"
    assert stored.importe == Decimal("1250.50")
    assert len(audit.events) == len(audit.changes) == 1
    assert uow.commits == 1


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"garantia_fecha_fin": date(2026, 10, 1)}, "garantia"),
        ({"importe": Decimal("-0.01")}, "importe"),
        ({"descripcion": "x" * 256}, "descripcion"),
    ],
)
async def test_rejects_invalid_milestone_without_side_effects(
    overrides: dict[str, object], message: str
) -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterHitoValidationError, match=message):
        await service.execute(_command(expediente_id, **overrides))

    assert not repo.by_id and not audit.events
    assert uow.commits == uow.rollbacks == 0


async def test_denies_missing_actor_before_writes() -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterHitoAuthorizationError, match="actor_id"):
        await service.execute(_command(expediente_id, actor_id=None))

    assert not repo.by_id and not audit.events
    assert uow.commits == 0


async def test_retry_with_same_id_and_payload_is_idempotent() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)

    first = await service.execute(command)
    second = await service.execute(command)

    assert second == first
    assert len(repo.by_id) == 1
    assert len(audit.events) == 1
    assert uow.commits == 1


async def test_reused_id_with_different_payload_conflicts() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)
    await service.execute(command)

    with pytest.raises(RegisterHitoConflictError, match="different payload"):
        await service.execute(
            _command(expediente_id, hito_id=command.hito_id, importe=Decimal("1"))
        )

    assert len(repo.by_id) == 1
    assert len(audit.events) == 1
    assert uow.commits == 1


async def test_dependency_failure_rolls_back_and_is_retryable() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)
    repo.fail = True

    with pytest.raises(RegisterHitoError, match="registration failed"):
        await service.execute(command)

    assert uow.rollbacks == 1
    assert not repo.by_id and not audit.events

    repo.fail = False
    result = await service.execute(command)
    assert result.hito_id == command.hito_id
    assert uow.commits == 1
