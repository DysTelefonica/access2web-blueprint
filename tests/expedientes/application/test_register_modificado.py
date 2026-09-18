"""Strict TDD for EXP-CAP-009 modification history."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.register_modificado.command import (
    RegisterModificadoAuthorizationError,
    RegisterModificadoCommand,
    RegisterModificadoConflictError,
    RegisterModificadoError,
    RegisterModificadoValidationError,
)
from app.src.modules.expedientes.application.register_modificado.service import (
    RegisterModificadoService,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


class _ExpedienteRepo:
    def __init__(self, expediente: Expediente) -> None:
        self.expediente = expediente

    async def get_by_id(self, expediente_id: UUID) -> object | None:
        return self.expediente if expediente_id == self.expediente.id else None


class _ModificadoRepo:
    def __init__(self) -> None:
        self.by_id: dict[UUID, Any] = {}
        self.fail = False

    async def get_by_id(self, modificado_id: UUID) -> object | None:
        return self.by_id.get(modificado_id)

    async def upsert(self, modificado: Any) -> object:
        if self.fail:
            raise RuntimeError("database unavailable")
        self.by_id[modificado.id] = modificado
        return modificado


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


def _make_service() -> tuple[RegisterModificadoService, _ModificadoRepo, _AuditLog, _UoW, UUID]:
    expediente = Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
    )
    repo = _ModificadoRepo()
    audit = _AuditLog()
    uow = _UoW()
    service = RegisterModificadoService(
        expediente_repo=_ExpedienteRepo(expediente),
        modificado_repo=repo,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    return service, repo, audit, uow, expediente.id


def _command(expediente_id: UUID, **overrides: Any) -> RegisterModificadoCommand:
    values: dict[str, Any] = {
        "modificado_id": uuid4(),
        "expediente_id": expediente_id,
        "n_modificado": "MOD-2026/01",
        "fecha_firma": date(2026, 6, 1),
        "fecha_fin": date(2026, 12, 1),
        "descripcion": "Prórroga del contrato",
        "actor_id": uuid4(),
    }
    values.update(overrides)
    return RegisterModificadoCommand(**values)


async def test_registers_auditable_modificado_atomically() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)

    result = await service.execute(command)

    stored = repo.by_id[command.modificado_id]
    assert result.modificado_id == command.modificado_id
    assert stored.id_expediente == expediente_id
    assert stored.n_modificado == "MOD-2026/01"
    assert stored.fecha_fin == date(2026, 12, 1)
    assert len(audit.events) == len(audit.changes) == 1
    assert uow.commits == 1


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"n_modificado": ""}, "n_modificado"),
        ({"n_modificado": "x" * 256}, "n_modificado"),
        ({"fecha_fin": date(2026, 1, 1)}, "fecha_fin"),
    ],
)
async def test_rejects_invalid_modificado_without_side_effects(
    overrides: dict[str, object], message: str
) -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterModificadoValidationError, match=message):
        await service.execute(_command(expediente_id, **overrides))

    assert not repo.by_id and not audit.events
    assert uow.commits == uow.rollbacks == 0


async def test_denies_missing_actor_before_writes() -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterModificadoAuthorizationError, match="actor_id"):
        await service.execute(_command(expediente_id, actor_id=None))

    assert not repo.by_id and not audit.events
    assert uow.commits == 0


async def test_retry_with_same_id_is_idempotent() -> None:
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

    with pytest.raises(RegisterModificadoConflictError, match="different payload"):
        await service.execute(
            _command(
                expediente_id,
                modificado_id=command.modificado_id,
                descripcion="Cambio menor",
            )
        )

    assert len(repo.by_id) == 1
    assert len(audit.events) == 1
    assert uow.commits == 1


async def test_dependency_failure_rolls_back_and_is_retryable() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)
    repo.fail = True

    with pytest.raises(RegisterModificadoError, match="modificado registration failed"):
        await service.execute(command)

    assert uow.rollbacks == 1
    assert not repo.by_id and not audit.events

    repo.fail = False
    result = await service.execute(command)
    assert result.modificado_id == command.modificado_id
    assert uow.commits == 1
