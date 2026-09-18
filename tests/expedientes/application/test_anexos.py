"""Strict TDD for EXP-CAP-010 anexos (create + delete)."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.anexos.command import (
    CreateAnexoAuthorizationError,
    CreateAnexoCommand,
    DeleteAnexoAuthorizationError,
    DeleteAnexoCommand,
    DeleteAnexoError,
    RegisterAnexoError,
    RegisterAnexoSizeLimitError,
    RegisterAnexoValidationError,
)
from app.src.modules.expedientes.application.anexos.service import AnexosService
from app.src.modules.expedientes.domain.anexo.reference import AnexoReference
from app.src.modules.expedientes.domain.anexo.retention import (
    RetentionLevel,
    RetentionPolicy,
)
from app.src.modules.expedientes.domain.anexo.size_limit import SizeLimitPolicy
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


class _ExpedienteRepo:
    def __init__(self, expediente: Expediente) -> None:
        self.expediente = expediente

    async def get_by_id(self, expediente_id: UUID) -> object | None:
        return self.expediente if expediente_id == self.expediente.id else None


class _AnexoRepo:
    def __init__(self) -> None:
        self.by_id: dict[UUID, Any] = {}
        self.deleted: list[UUID] = []
        self.fail = False

    async def get_by_id(self, anexo_id: UUID) -> object | None:
        return self.by_id.get(anexo_id)

    async def get_by_expediente(self, expediente_id: UUID) -> list[object]:
        return [a for a in self.by_id.values() if a.id_expediente == expediente_id]

    async def add(self, anexo: Any) -> object:
        if self.fail:
            raise RuntimeError("database unavailable")
        self.by_id[anexo.id] = anexo
        return anexo

    async def delete(self, anexo_id: UUID) -> None:
        if self.fail:
            raise RuntimeError("database unavailable")
        self.deleted.append(anexo_id)
        self.by_id.pop(anexo_id, None)


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


def _expediente() -> Expediente:
    return Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
    )


def _reference(size: int = 1024) -> AnexoReference:
    return AnexoReference(
        storage_ref=f"storage://exp/{uuid4()}",
        content_type="application/pdf",
        size_bytes=size,
    )


def _make_service() -> tuple[
    AnexosService, _AnexoRepo, _AuditLog, _UoW, Expediente, SizeLimitPolicy
]:
    expediente = _expediente()
    anexo_repo = _AnexoRepo()
    audit = _AuditLog()
    uow = _UoW()
    size_limit = SizeLimitPolicy(max_size_bytes=10 * 1024)
    service = AnexosService(
        expediente_repo=_ExpedienteRepo(expediente),
        anexo_repo=anexo_repo,
        audit_log=audit,
        size_limit=size_limit,
        uow_factory=lambda: uow,
    )
    return service, anexo_repo, audit, uow, expediente, size_limit


def _create_command(expediente_id: UUID, **overrides: Any) -> CreateAnexoCommand:
    values: dict[str, Any] = {
        "anexo_id": uuid4(),
        "expediente_id": expediente_id,
        "nombre": "informe.pdf",
        "referencia": _reference(2048),
        "retention": RetentionPolicy.legal_default(),
        "actor_id": uuid4(),
    }
    values.update(overrides)
    return CreateAnexoCommand(**values)


async def test_create_persists_auditable_anexo_atomically() -> None:
    service, repo, audit, uow, expediente, _ = _make_service()
    command = _create_command(expediente.id)

    result = await service.create(command)

    stored = repo.by_id[command.anexo_id]
    assert result.anexo_id == command.anexo_id
    assert stored.id_expediente == expediente.id
    assert stored.referencia.storage_ref == command.referencia.storage_ref
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "anexo.created"
    assert len(audit.changes) == 1
    assert audit.changes[0].accion == "alta"
    assert uow.commits == 1


async def test_create_rejects_missing_actor_before_writes() -> None:
    service, repo, audit, uow, expediente, _ = _make_service()

    with pytest.raises(CreateAnexoAuthorizationError, match="actor_id"):
        await service.create(_create_command(expediente.id, actor_id=None))

    assert not repo.by_id
    assert not audit.events
    assert uow.commits == uow.rollbacks == 0


async def test_create_rejects_size_over_limit() -> None:
    service, repo, audit, uow, expediente, _ = _make_service()
    over_ref = _reference(size=11 * 1024)

    with pytest.raises(RegisterAnexoSizeLimitError, match="limit"):
        await service.create(_create_command(expediente.id, referencia=over_ref))

    assert not repo.by_id
    assert not audit.events
    assert uow.commits == uow.rollbacks == 0


async def test_create_rejects_unknown_expediente() -> None:
    service, repo, audit, uow, _, _ = _make_service()

    with pytest.raises(RegisterAnexoValidationError, match="not found"):
        await service.create(_create_command(uuid4()))

    assert not repo.by_id and not audit.events


async def test_dependency_failure_on_create_rolls_back() -> None:
    service, repo, audit, uow, expediente, _ = _make_service()
    command = _create_command(expediente.id)
    repo.fail = True

    with pytest.raises(RegisterAnexoError):
        await service.create(command)

    assert uow.rollbacks == 1
    assert not repo.by_id and not audit.events


async def test_delete_removes_anexo_and_records_audit_event() -> None:
    service, repo, audit, uow, expediente, _ = _make_service()
    create_cmd = _create_command(expediente.id)
    await service.create(create_cmd)
    before = datetime.now(UTC)

    delete_cmd = DeleteAnexoCommand(
        anexo_id=create_cmd.anexo_id,
        expected_version=1,
        actor_id=create_cmd.actor_id,
    )
    result = await service.delete(delete_cmd)

    after = datetime.now(UTC)
    assert result.anexo_id == create_cmd.anexo_id
    assert before <= result.deleted_at <= after
    assert create_cmd.anexo_id in repo.deleted
    assert audit.events[-1].event_type == "anexo.deleted"


async def test_delete_rejects_missing_actor() -> None:
    service, repo, audit, uow, expediente, _ = _make_service()
    create_cmd = _create_command(expediente.id)
    await service.create(create_cmd)

    with pytest.raises(DeleteAnexoAuthorizationError):
        await service.delete(
            DeleteAnexoCommand(
                anexo_id=create_cmd.anexo_id,
                expected_version=1,
                actor_id=None,
            )
        )

    assert create_cmd.anexo_id in repo.by_id
    assert uow.commits == 1  # only the create transaction committed


async def test_delete_rejects_unknown_anexo() -> None:
    service, _, _, uow, _, _ = _make_service()

    with pytest.raises(RegisterAnexoValidationError, match="not found"):
        await service.delete(
            DeleteAnexoCommand(
                anexo_id=uuid4(),
                expected_version=1,
                actor_id=uuid4(),
            )
        )

    assert uow.commits == 0


async def test_delete_blocks_expired_retention() -> None:
    """Anexo with retention already expired at delete time cannot be
    deleted (CAP-010 §Concurrencia: preserve evidence)."""
    service, repo, audit, uow, expediente, _ = _make_service()
    create_cmd = _create_command(expediente.id, retention=RetentionPolicy.historical())
    await service.create(create_cmd)
    stored = repo.by_id[create_cmd.anexo_id]
    stored.created_at = datetime.now(UTC) - timedelta(days=400)
    stored.retention = RetentionPolicy(
        level=RetentionLevel.OPERATIONAL, duration=timedelta(days=30)
    )

    with pytest.raises(DeleteAnexoError, match="retention expired"):
        await service.delete(
            DeleteAnexoCommand(
                anexo_id=create_cmd.anexo_id,
                expected_version=1,
                actor_id=create_cmd.actor_id,
            )
        )

    assert create_cmd.anexo_id in repo.by_id
    assert audit.events[-1].event_type == "anexo.created"


async def test_dependency_failure_on_delete_rolls_back() -> None:
    service, repo, audit, uow, expediente, _ = _make_service()
    create_cmd = _create_command(expediente.id)
    await service.create(create_cmd)
    delete_cmd = DeleteAnexoCommand(
        anexo_id=create_cmd.anexo_id,
        expected_version=1,
        actor_id=create_cmd.actor_id,
    )
    repo.fail = True

    with pytest.raises(DeleteAnexoError):
        await service.delete(delete_cmd)

    assert create_cmd.anexo_id in repo.by_id
    assert uow.commits == 1
