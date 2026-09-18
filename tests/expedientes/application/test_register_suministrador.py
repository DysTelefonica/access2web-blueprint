"""Strict TDD for EXP-CAP-014 suministrador + UTE hierarchy."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.register_suministrador.command import (
    RegisterSuministradorAuthorizationError,
    RegisterSuministradorCommand,
    RegisterSuministradorConflictError,
    RegisterSuministradorError,
    RegisterSuministradorValidationError,
)
from app.src.modules.expedientes.application.register_suministrador.service import (
    RegisterSuministradorService,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


class _ExpedienteRepo:
    def __init__(self, expediente: Expediente) -> None:
        self.expediente = expediente

    async def get_by_id(self, expediente_id: UUID) -> object | None:
        return self.expediente if expediente_id == self.expediente.id else None


class _SuministradorRepo:
    def __init__(self) -> None:
        self.by_id: dict[UUID, Any] = {}
        self.fail = False

    async def get_by_id(self, suministrador_id: UUID) -> object | None:
        return self.by_id.get(suministrador_id)

    async def upsert(self, suministrador: Any) -> object:
        if self.fail:
            raise RuntimeError("database unavailable")
        self.by_id[suministrador.id] = suministrador
        return suministrador


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


def _make_service() -> tuple[
    RegisterSuministradorService, _SuministradorRepo, _AuditLog, _UoW, UUID
]:
    expediente = Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
    )
    repo = _SuministradorRepo()
    audit = _AuditLog()
    uow = _UoW()
    service = RegisterSuministradorService(
        expediente_repo=_ExpedienteRepo(expediente),
        suministrador_repo=repo,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    return service, repo, audit, uow, expediente.id


def _command(expediente_id: UUID, **overrides: Any) -> RegisterSuministradorCommand:
    values: dict[str, Any] = {
        "suministrador_id": uuid4(),
        "expediente_id": expediente_id,
        "id_suministrador": uuid4(),
        "id_padre": None,
        "descripcion": "UTE principal",
        "contratista_principal": True,
        "sub_contratista": False,
        "actor_id": uuid4(),
    }
    values.update(overrides)
    return RegisterSuministradorCommand(**values)


async def test_registers_auditable_root_suministrador_atomically() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)

    result = await service.execute(command)

    stored = repo.by_id[command.suministrador_id]
    assert result.suministrador_id == command.suministrador_id
    assert stored.id_expediente == expediente_id
    assert stored.id_padre is None
    assert stored.descripcion == "UTE principal"
    assert stored.contratista_principal is True
    assert stored.sub_contratista is False
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "suministrador.registered"
    assert len(audit.changes) == 1
    assert audit.changes[0].accion == "alta"
    assert uow.commits == 1


async def test_root_requires_contratista_principal_flag() -> None:
    """The CAP-014 rule «el árbol manda» requires a single root per
    expediente with ``contratista_principal=True``. A root row that
    omits the flag is rejected so the hierarchy stays deterministic.
    """
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterSuministradorValidationError, match="contratista"):
        await service.execute(_command(expediente_id, contratista_principal=False))

    assert not repo.by_id and not audit.events
    assert uow.commits == uow.rollbacks == 0


async def test_sub_contratista_requires_parent_in_same_expediente() -> None:
    """Subcontratistas inherit from a parent in the same expediente.

    The parent must exist (and belong to the same expediente); otherwise
    the service rejects the row without side effects.
    """
    service, repo, audit, uow, expediente_id = _make_service()
    orphan_parent = uuid4()

    with pytest.raises(RegisterSuministradorValidationError, match="parent suministrador"):
        await service.execute(
            _command(
                expediente_id,
                sub_contratista=True,
                contratista_principal=True,
                id_padre=orphan_parent,
            )
        )

    assert not repo.by_id and not audit.events
    assert uow.commits == uow.rollbacks == 0


async def test_parent_loop_is_rejected() -> None:
    """A parent that resolves to the same row would form a 1-cycle. The
    CAP-014 «el árbol manda» invariant forbids cycles."""
    service, repo, audit, uow, expediente_id = _make_service()

    parent_row_id = uuid4()
    command = _command(
        expediente_id,
        id_suministrador=parent_row_id,
        contratista_principal=True,
        sub_contratista=False,
    )
    await service.execute(command)

    cycle = _command(
        expediente_id,
        suministrador_id=parent_row_id,
        id_padre=parent_row_id,
        contratista_principal=True,
        sub_contratista=True,
    )

    with pytest.raises(RegisterSuministradorValidationError, match="cycle"):
        await service.execute(cycle)

    assert cycle.suministrador_id not in repo.by_id
    assert uow.commits == 1


async def test_denies_missing_actor_before_writes() -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterSuministradorAuthorizationError, match="actor_id"):
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

    with pytest.raises(RegisterSuministradorConflictError, match="different payload"):
        await service.execute(
            _command(
                expediente_id,
                suministrador_id=command.suministrador_id,
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

    with pytest.raises(RegisterSuministradorError, match="suministrador registration failed"):
        await service.execute(command)

    assert uow.rollbacks == 1
    assert not repo.by_id and not audit.events

    repo.fail = False
    result = await service.execute(command)
    assert result.suministrador_id == command.suministrador_id
    assert uow.commits == 1
