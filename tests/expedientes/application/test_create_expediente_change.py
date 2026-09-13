"""Strict TDD — ExpedienteAltaService records the change (issue #668).

C01 (#669) implements CAP-001 §Camino feliz partially: cabecera + initial
hito + audit event. Issue #668 closes the gap by adding:

- ``TbUltimoCambio`` row (the last-change marker).
- ``TbCambios`` row (the historical log).

Both written inside the same UoW so a failure rolls back the cabecera
too (DA-11 audit-in-same-tx).
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any
from uuid import uuid4

import pytest

from app.src.modules.expedientes.application.create_expediente.command import (
    ExpedienteAltaCommand,
)
from app.src.modules.expedientes.application.create_expediente.service import (
    ExpedienteAltaService,
)
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo
from app.src.modules.expedientes.ports.audit_log import (
    AuditLogPort,
    ChangeRecord,
    ExpedienteAuditEvent,
)
from app.src.modules.expedientes.ports.expediente_repository import (
    ExpedienteRepositoryPort,
)
from app.src.modules.expedientes.ports.hito_repository import HitoRepositoryPort

# ---------------------------------------------------------------------------
# Doubles
# ---------------------------------------------------------------------------


class _FakeExpedienteRepository(ExpedienteRepositoryPort):
    def __init__(self) -> None:
        self.by_id: dict[Any, object] = {}
        self.calls: list[str] = []

    async def get_by_id(self, expediente_id: Any) -> object | None:
        self.calls.append("get_by_id")
        return self.by_id.get(expediente_id)

    async def create(self, aggregate: object) -> object:
        self.calls.append("create")
        self.by_id[aggregate.id] = aggregate
        return aggregate

    async def update(self, aggregate: object) -> object:
        self.calls.append("update")
        self.by_id[aggregate.id] = aggregate
        return aggregate

    async def delete(self, expediente_id: Any) -> None:
        self.calls.append("delete")
        self.by_id.pop(expediente_id, None)

    async def list_by_state(self, estado: str, limit: int, offset: int) -> tuple[list[object], int]:
        return [], 0


class _FakeHitoRepository(HitoRepositoryPort):
    def __init__(self) -> None:
        self.by_exp: dict[Any, list[object]] = {}
        self.calls: list[str] = []

    async def get_by_expediente(self, expediente_id: Any) -> list[object]:
        return list(self.by_exp.get(expediente_id, []))

    async def upsert(self, hito: object) -> object:
        self.calls.append("upsert")
        eid = getattr(hito, "id_expediente", None)
        if eid is not None:
            self.by_exp.setdefault(eid, []).append(hito)
        return hito

    async def delete(self, hito_id: Any) -> None:
        pass


class _FakeAuditLog(AuditLogPort):
    def __init__(self) -> None:
        self.events: list[ExpedienteAuditEvent] = []
        self.changes: list[ChangeRecord] = []
        self.calls: list[str] = []

    async def append(self, event: ExpedienteAuditEvent) -> None:
        self.calls.append("append")
        self.events.append(event)

    async def list_for_actor(self, actor_id: Any, since: datetime) -> list[ExpedienteAuditEvent]:
        return []

    async def record_change(self, change: ChangeRecord) -> None:
        self.calls.append("record_change")
        self.changes.append(change)


class _FakeUoW:
    def __init__(self, session: object) -> None:
        self._session = session
        self.commits = 0
        self.rollbacks = 0

    def __enter__(self) -> object:
        return self._session

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        if exc_type is None:
            self.commits += 1
        else:
            self.rollbacks += 1


class _FakeSession:
    pass


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_command(**overrides: Any) -> ExpedienteAltaCommand:
    defaults: dict[str, Any] = {
        "codigo_ordinal": "EXP-001",
        "tipo": ExpedienteTipo.AM,
        "actor_id": uuid4(),
    }
    defaults.update(overrides)
    return ExpedienteAltaCommand(**defaults)


def _make_service() -> tuple[
    ExpedienteAltaService, _FakeUoW, _FakeExpedienteRepository, _FakeHitoRepository, _FakeAuditLog
]:
    repo = _FakeExpedienteRepository()
    hito_repo = _FakeHitoRepository()
    audit = _FakeAuditLog()
    uow = _FakeUoW(_FakeSession())
    service = ExpedienteAltaService(
        expediente_repo=repo,
        hito_repo=hito_repo,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    return service, uow, repo, hito_repo, audit


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


async def test_alta_records_change_in_audit_log() -> None:
    """§Camino feliz: the change row is written inside the UoW."""
    service, _, _, _, audit = _make_service()
    cmd = _make_command()

    await service.execute(cmd)

    assert "record_change" in audit.calls
    assert len(audit.changes) == 1


async def test_change_record_carries_alta_action() -> None:
    service, _, _, _, audit = _make_service()
    cmd = _make_command()

    await service.execute(cmd)

    assert audit.changes[0].accion == "alta"
    assert audit.changes[0].nombre_tabla == "expedientes"


async def test_change_record_carries_actor_and_expediente() -> None:
    service, _, _, _, audit = _make_service()
    actor = uuid4()
    cmd = _make_command(actor_id=actor)

    result = await service.execute(cmd)

    change = audit.changes[0]
    assert change.id_usuario_cambio == actor
    assert change.id_expediente == result.expediente_id


async def test_change_record_carries_creation_timestamp() -> None:
    service, _, _, _, audit = _make_service()
    before = datetime.now(UTC)

    await service.execute(_make_command())

    after = datetime.now(UTC)
    change = audit.changes[0]
    assert before <= change.fecha_cambio <= after


async def test_change_record_is_rolled_back_on_failure() -> None:
    """§Concurrencia o fallo: if the cabecera is rolled back, the
    change row goes with it. The UoW is the unit of atomicity.
    """

    class _BrokenRepo(_FakeExpedienteRepository):
        async def create(self, aggregate: object) -> object:
            raise RuntimeError("simulated adapter failure")

    repo = _BrokenRepo()
    hito_repo = _FakeHitoRepository()
    audit = _FakeAuditLog()
    uow = _FakeUoW(_FakeSession())
    service = ExpedienteAltaService(
        expediente_repo=repo,
        hito_repo=hito_repo,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    cmd = _make_command()

    from app.src.modules.expedientes.application.create_expediente.command import (
        ExpedienteAltaError,
    )

    with pytest.raises(ExpedienteAltaError):
        await service.execute(cmd)

    assert audit.changes == []


async def test_change_record_does_not_leak_after_successful_alta() -> None:
    """After a successful alta, the change record exists; before any
    alta runs, the fake is empty. This guards against accidentally
    reusing state across tests.
    """
    service, _, _, _, audit = _make_service()

    assert audit.changes == []
    await service.execute(_make_command())
    assert len(audit.changes) == 1
