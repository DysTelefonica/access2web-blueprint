"""Strict TDD for EXP-CAP-011 anualidades."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.register_anualidad.command import (
    RegisterAnualidadAuthorizationError,
    RegisterAnualidadCommand,
    RegisterAnualidadConflictError,
    RegisterAnualidadError,
    RegisterAnualidadValidationError,
)
from app.src.modules.expedientes.application.register_anualidad.service import (
    RegisterAnualidadService,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


class _ExpedienteRepo:
    def __init__(self, expediente: Expediente) -> None:
        self.expediente = expediente

    async def get_by_id(self, expediente_id: UUID) -> object | None:
        return self.expediente if expediente_id == self.expediente.id else None


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


def _make_service() -> tuple[RegisterAnualidadService, _AnualidadRepo, _AuditLog, _UoW, UUID]:
    expediente = Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
    )
    repo = _AnualidadRepo()
    audit = _AuditLog()
    uow = _UoW()
    service = RegisterAnualidadService(
        expediente_repo=_ExpedienteRepo(expediente),
        anualidad_repo=repo,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    return service, repo, audit, uow, expediente.id


def _command(expediente_id: UUID, **overrides: Any) -> RegisterAnualidadCommand:
    values: dict[str, Any] = {
        "anualidad_id": uuid4(),
        "expediente_id": expediente_id,
        "anio": 2026,
        "bi_iva": Decimal("1000.00"),
        "bi_ipsi": None,
        "bi_igic": None,
        "bi_exenta": Decimal("0.00"),
        "iva": Decimal("210.00"),
        "ipsi": None,
        "igic": None,
        "periodo_facturacion": "Q1",
        "actor_id": uuid4(),
    }
    values.update(overrides)
    return RegisterAnualidadCommand(**values)


async def test_registers_auditable_anualidad_atomically() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)

    result = await service.execute(command)

    stored = repo.by_id[command.anualidad_id]
    assert result.anualidad_id == command.anualidad_id
    assert stored.id_expediente == expediente_id
    assert stored.anio == 2026
    assert stored.bi_iva == Decimal("1000.00")
    assert stored.iva == Decimal("210.00")
    assert stored.periodo_facturacion == "Q1"
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "anualidad.registered"
    assert len(audit.changes) == 1
    assert audit.changes[0].accion == "alta"
    assert uow.commits == 1


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"anio": 1969}, "anio"),
        ({"anio": 2200}, "anio"),
        ({"iva": Decimal("-0.01")}, "iva"),
        ({"bi_iva": Decimal("-1.00")}, "bi_iva"),
        ({"periodo_facturacion": "x" * 256}, "periodo_facturacion"),
    ],
)
async def test_rejects_invalid_anualidad_without_side_effects(
    overrides: dict[str, object], message: str
) -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterAnualidadValidationError, match=message):
        await service.execute(_command(expediente_id, **overrides))

    assert not repo.by_id and not audit.events
    assert uow.commits == uow.rollbacks == 0


async def test_denies_missing_actor_before_writes() -> None:
    service, repo, audit, uow, expediente_id = _make_service()

    with pytest.raises(RegisterAnualidadAuthorizationError, match="actor_id"):
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

    with pytest.raises(RegisterAnualidadConflictError, match="different payload"):
        await service.execute(
            _command(
                expediente_id,
                anualidad_id=command.anualidad_id,
                iva=Decimal("250.00"),
            )
        )

    assert len(repo.by_id) == 1
    assert len(audit.events) == 1
    assert uow.commits == 1


async def test_dependency_failure_rolls_back_and_is_retryable() -> None:
    service, repo, audit, uow, expediente_id = _make_service()
    command = _command(expediente_id)
    repo.fail = True

    with pytest.raises(RegisterAnualidadError, match="anualidad registration failed"):
        await service.execute(command)

    assert uow.rollbacks == 1
    assert not repo.by_id and not audit.events

    repo.fail = False
    result = await service.execute(command)
    assert result.anualidad_id == command.anualidad_id
    assert uow.commits == 1
