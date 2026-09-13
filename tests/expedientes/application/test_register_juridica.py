"""Strict TDD for EXP-CAP-013 entidad/juridica relations."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.register_juridica.command import (
    RegisterJuridicaAuthorizationError,
    RegisterJuridicaCommand,
    RegisterJuridicaConflictError,
    RegisterJuridicaError,
    RegisterJuridicaValidationError,
)
from app.src.modules.expedientes.application.register_juridica.service import (
    RegisterJuridicaService,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


class _ExpedienteRepo:
    def __init__(self, expediente: Expediente) -> None:
        self.expediente = expediente

    async def get_by_id(self, expediente_id: UUID) -> object | None:
        return self.expediente if expediente_id == self.expediente.id else None


class _JuridicaRepo:
    def __init__(self) -> None:
        self.by_id: dict[UUID, Any] = {}
        self.fail = False

    async def get_by_id(self, juridica_id: UUID) -> object | None:
        return self.by_id.get(juridica_id)

    async def upsert(self, juridica: Any) -> object:
        if self.fail:
            raise RuntimeError("database unavailable")
        self.by_id[juridica.id] = juridica
        return juridica


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


def _make_service() -> tuple[RegisterJuridicaService, _JuridicaRepo, _AuditLog, _UoW, UUID]:
    expediente = Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
    )
    repo = _JuridicaRepo()
    audit = _AuditLog()
    uow = _UoW()
    service = RegisterJuridicaService(
        expediente_repo=_ExpedienteRepo(expediente),
        juridica_repo=repo,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    return service, repo, audit, uow, expediente.id


def _command(expediente_id: UUID, **overrides: Any) -> RegisterJuridicaCommand:
    values: dict[str, Any] = {
        "juridica_id": uuid4(),
        "expediente_id": expediente_id,
        "id_juridica": uuid4(),
        "id_suministrador": None,
        "contratista_principal": True,
        "sub_contratista": False,
        "actor_id": uuid4(),
    }
    values.update(overrides)
    return RegisterJuridicaCommand(**values)


async def test_registers_auditable_juridica_atomically() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)

    result = await service.execute(command)

    stored = repo.by_id[command.juridica_id]
    assert result.juridica_id == command.juridica_id
    assert stored.id_expediente == expediente_id
    assert stored.id_juridica == command.id_juridica
    assert stored.contratista_principal is True
    assert stored.sub_contratista is False
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "juridica.registered"
    assert len(audit.changes) == 1
    assert audit.changes[0].accion == "alta"
    assert uow.commits == 1


async def test_denies_missing_actor_before_writes() -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterJuridicaAuthorizationError, match="actor_id"):
        await service.execute(_command(expediente_id, actor_id=None))

    assert not repo.by_id and not audit.events
    assert uow.commits == 0


async def test_rejects_missing_juridica_reference() -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterJuridicaValidationError, match="id_juridica"):
        await service.execute(_command(expediente_id, id_juridica=None))

    assert not repo.by_id and not audit.events
    assert uow.commits == uow.rollbacks == 0


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

    with pytest.raises(RegisterJuridicaConflictError, match="different payload"):
        await service.execute(
            _command(
                expediente_id,
                juridica_id=command.juridica_id,
                sub_contratista=True,
            )
        )

    assert len(repo.by_id) == 1
    assert len(audit.events) == 1
    assert uow.commits == 1


async def test_dependency_failure_rolls_back_and_is_retryable() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)
    repo.fail = True

    with pytest.raises(RegisterJuridicaError, match="juridica registration failed"):
        await service.execute(command)

    assert uow.rollbacks == 1
    assert not repo.by_id and not audit.events

    repo.fail = False
    result = await service.execute(command)
    assert result.juridica_id == command.juridica_id
    assert uow.commits == 1
