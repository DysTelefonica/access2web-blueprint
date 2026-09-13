"""Strict TDD for EXP-CAP-031 autosave of related-data sections."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.autosave_related.command import (
    AutosaveRelatedAuthorizationError,
    AutosaveRelatedCommand,
    AutosaveRelatedConflictError,
    AutosaveRelatedError,
    AutosaveRelatedValidationError,
)
from app.src.modules.expedientes.application.autosave_related.service import (
    AutosaveRelatedService,
)
from app.src.modules.expedientes.domain.anualidad import Anualidad
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo
from app.src.modules.expedientes.domain.hito import Hito
from app.src.modules.expedientes.domain.modificado import Modificado


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


class _AnualidadRepo:
    def __init__(self) -> None:
        self.by_id: dict[UUID, Any] = {}
        self.fail = False

    async def get_by_id(self, anualidad_id: UUID) -> object | None:
        return self.by_id.get(anualidad_id)

    async def upsert(self, anualidad: Any) -> object:
        if self.fail:
            raise RuntimeError("database unavailable")
        self.by_id[anualidad.id] = anualidad
        return anualidad


class _Idempotency:
    def __init__(self) -> None:
        self.records: dict[UUID, Any] = {}

    async def get(self, key: UUID) -> object | None:
        return self.records.get(key)

    async def record(self, record: Any) -> None:
        self.records[record.key] = record


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
        version=3,
    )


def _make_service() -> tuple[
    AutosaveRelatedService,
    _HitoRepo,
    _ModificadoRepo,
    _AnualidadRepo,
    _Idempotency,
    _AuditLog,
    _UoW,
    UUID,
]:
    expediente = _expediente()
    hito_repo = _HitoRepo()
    mod_repo = _ModificadoRepo()
    anu_repo = _AnualidadRepo()
    idemp = _Idempotency()
    audit = _AuditLog()
    uow = _UoW()
    service = AutosaveRelatedService(
        expediente_repo=_ExpedienteRepo(expediente),
        hito_repo=hito_repo,
        modificado_repo=mod_repo,
        anualidad_repo=anu_repo,
        idempotency=idemp,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    return service, hito_repo, mod_repo, anu_repo, idemp, audit, uow, expediente.id


def _hito(expediente_id: UUID, *, hito_id: UUID | None = None) -> Hito:
    return Hito(
        id=hito_id or uuid4(),
        id_expediente=expediente_id,
        fecha_hito=date(2026, 10, 1),
        garantia_fecha_fin=date(2027, 10, 1),
        estado=ExpedienteEstado.BORRADOR,
    )


def _modificado(expediente_id: UUID, *, modificado_id: UUID | None = None) -> Modificado:
    return Modificado(
        id=modificado_id or uuid4(),
        id_expediente=expediente_id,
        n_modificado="MOD-2026/01",
        fecha_firma=date(2026, 6, 1),
        fecha_fin=date(2026, 12, 1),
        descripcion="Prórroga",
    )


def _anualidad(expediente_id: UUID, *, anualidad_id: UUID | None = None) -> Anualidad:
    return Anualidad(
        id=anualidad_id or uuid4(),
        id_expediente=expediente_id,
        anio=2026,
        bi_iva=Decimal("1000.00"),
        bi_ipsi=None,
        bi_igic=None,
        bi_exenta=Decimal("0.00"),
        iva=Decimal("210.00"),
        ipsi=None,
        igic=None,
        periodo_facturacion="Q1",
    )


def _command(expediente_id: UUID, **overrides: Any) -> AutosaveRelatedCommand:
    values: dict[str, Any] = {
        "idempotency_key": uuid4(),
        "expediente_id": expediente_id,
        "expected_version": 3,
        "hitos": [_hito(expediente_id)],
        "modificados": [_modificado(expediente_id)],
        "anualidades": [_anualidad(expediente_id)],
        "actor_id": uuid4(),
    }
    values.update(overrides)
    return AutosaveRelatedCommand(**values)


async def test_autosave_related_persists_hitos_modificados_anualidades() -> None:
    (
        service,
        hito_repo,
        mod_repo,
        anu_repo,
        idemp,
        audit,
        uow,
        expediente_id,
    ) = _make_service()
    command = _command(expediente_id)

    result = await service.execute(command)

    assert result.hitos_added == 1
    assert result.modificados_added == 1
    assert result.anualidades_added == 1
    assert len(hito_repo.by_id) == 1
    assert len(mod_repo.by_id) == 1
    assert len(anu_repo.by_id) == 1
    assert uow.commits == 1
    assert command.idempotency_key in idemp.records
    assert len(audit.events) == 3
    assert len(audit.changes) == 3


async def test_autosave_related_rejects_stale_version() -> None:
    (
        service,
        _,
        _,
        _,
        _,
        _,
        _,
        expediente_id,
    ) = _make_service()
    command = _command(expediente_id, expected_version=2)

    with pytest.raises(AutosaveRelatedConflictError, match="version stale"):
        await service.execute(command)


async def test_autosave_related_idempotent_replay_returns_recorded_result() -> None:
    (
        service,
        _,
        _,
        _,
        _,
        audit,
        uow,
        expediente_id,
    ) = _make_service()
    command = _command(expediente_id)

    first = await service.execute(command)
    second = await service.execute(command)

    assert second == first
    assert len(audit.events) == 3
    assert uow.commits == 1


async def test_autosave_related_rejects_idempotency_reuse_with_different_payload() -> None:
    (
        service,
        _,
        _,
        _,
        _,
        audit,
        uow,
        expediente_id,
    ) = _make_service()
    command = _command(expediente_id)

    await service.execute(command)

    different = _command(expediente_id)
    different = AutosaveRelatedCommand(
        idempotency_key=command.idempotency_key,
        expediente_id=command.expediente_id,
        expected_version=command.expected_version,
        hitos=[],
        modificados=command.modificados,
        anualidades=command.anualidades,
        actor_id=command.actor_id,
    )

    with pytest.raises(AutosaveRelatedConflictError, match="different payload"):
        await service.execute(different)

    assert uow.commits == 1
    assert len(audit.events) == 3


async def test_autosave_related_denies_missing_actor() -> None:
    service, *_ = _make_service()

    with pytest.raises(AutosaveRelatedAuthorizationError, match="actor_id"):
        await service.execute(_command(_expediente().id, actor_id=None))


async def test_autosave_related_dependency_failure_rolls_back() -> None:
    (
        service,
        _,
        _,
        anu_repo,
        _,
        audit,
        uow,
        expediente_id,
    ) = _make_service()
    command = _command(expediente_id)
    anu_repo.fail = True

    with pytest.raises(AutosaveRelatedError, match="related autosave failed"):
        await service.execute(command)

    assert uow.rollbacks == 1
    assert not audit.events


async def test_autosave_related_rejects_invalid_modificado_payload() -> None:
    service, _, _, _, _, _, _, expediente_id = _make_service()
    bad_mod = _modificado(expediente_id)
    bad_mod.fecha_fin = date(2026, 1, 1)
    command = _command(expediente_id, modificados=[bad_mod])

    with pytest.raises(AutosaveRelatedValidationError, match="fecha_fin"):
        await service.execute(command)
