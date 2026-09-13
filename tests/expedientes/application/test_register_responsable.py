"""Strict TDD for EXP-CAP-012 responsable assignment."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.register_responsable.command import (
    RegisterResponsableAuthorizationError,
    RegisterResponsableCommand,
    RegisterResponsableConflictError,
    RegisterResponsableError,
    RegisterResponsableValidationError,
)
from app.src.modules.expedientes.application.register_responsable.service import (
    RegisterResponsableService,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


class _ExpedienteRepo:
    def __init__(self, expediente: Expediente) -> None:
        self.expediente = expediente

    async def get_by_id(self, expediente_id: UUID) -> object | None:
        return self.expediente if expediente_id == self.expediente.id else None


class _ResponsableRepo:
    def __init__(self) -> None:
        self.by_id: dict[UUID, Any] = {}
        self.fail = False

    async def get_by_id(self, responsable_id: UUID) -> object | None:
        return self.by_id.get(responsable_id)

    async def upsert(self, responsable: Any) -> object:
        if self.fail:
            raise RuntimeError("database unavailable")
        self.by_id[responsable.id] = responsable
        return responsable


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


def _make_service() -> tuple[RegisterResponsableService, _ResponsableRepo, _AuditLog, _UoW, UUID]:
    expediente = Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
    )
    repo = _ResponsableRepo()
    audit = _AuditLog()
    uow = _UoW()
    service = RegisterResponsableService(
        expediente_repo=_ExpedienteRepo(expediente),
        responsable_repo=repo,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    return service, repo, audit, uow, expediente.id


def _command(expediente_id: UUID, **overrides: Any) -> RegisterResponsableCommand:
    values: dict[str, Any] = {
        "responsable_id": uuid4(),
        "expediente_id": expediente_id,
        "usuario_id": uuid4(),
        "rol": "JP",
        "correo_siempre": True,
        "es_jefe_proyecto": True,
        "es_preventa": False,
        "actor_id": uuid4(),
    }
    values.update(overrides)
    return RegisterResponsableCommand(**values)


async def test_registers_auditable_responsable_atomically() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)

    result = await service.execute(command)

    stored = repo.by_id[command.responsable_id]
    assert result.responsable_id == command.responsable_id
    assert stored.id_expediente == expediente_id
    assert stored.rol == "JP"
    assert stored.es_jefe_proyecto is True
    assert stored.correo_siempre is True
    assert stored.es_preventa is False
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "responsable.registered"
    assert len(audit.changes) == 1
    assert audit.changes[0].accion == "alta"
    assert uow.commits == 1


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"rol": ""}, "rol"),
        ({"rol": "x" * 65}, "rol"),
    ],
)
async def test_rejects_invalid_responsable_without_side_effects(
    overrides: dict[str, object], message: str
) -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterResponsableValidationError, match=message):
        await service.execute(_command(expediente_id, **overrides))

    assert not repo.by_id and not audit.events
    assert uow.commits == uow.rollbacks == 0


async def test_denies_missing_actor_before_writes() -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterResponsableAuthorizationError, match="actor_id"):
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

    with pytest.raises(RegisterResponsableConflictError, match="different payload"):
        await service.execute(
            _command(
                expediente_id,
                responsable_id=command.responsable_id,
                es_preventa=True,
            )
        )

    assert len(repo.by_id) == 1
    assert len(audit.events) == 1
    assert uow.commits == 1


async def test_dependency_failure_rolls_back_and_is_retryable() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)
    repo.fail = True

    with pytest.raises(RegisterResponsableError, match="responsable registration failed"):
        await service.execute(command)

    assert uow.rollbacks == 1
    assert not repo.by_id and not audit.events

    repo.fail = False
    result = await service.execute(command)
    assert result.responsable_id == command.responsable_id
    assert uow.commits == 1
