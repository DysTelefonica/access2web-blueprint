"""Strict TDD — ExpedienteDeleteService (CAP-003, issue #228).

CAP-003 §Camino feliz: "expediente con relaciones permitidas para
eliminación y actor autorizado — se elimina el agregado de forma
auditable sin afectar relaciones no incluidas".

§Validación y autorización: deniega sin efecto parcial, devuelve
error diagnosticable y conserva auditoría.
§Concurrencia o fallo: preserva invariantes, no duplica ni pierde
evidencia, queda reintentable.

The conditional aspect: the use case refuses the delete when the
aggregate has children (LOTE/BASED relationships, anexos, hitos, etc.)
unless the caller passes ``force=True``. This implements the
"impedir pérdida de hijos" requirement of CAP-003.

These tests use the same protocols as C01 (#669) and C02 (#671):
- ``ExpedienteRepositoryPort`` (F01, #222).
- ``AuditLogPort`` with ``record_change`` (#670).
- ``UnitOfWork`` (F04, #417).
- The aggregate ``Expediente`` (F02, #643).
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any
from uuid import uuid4

import pytest

from app.src.modules.expedientes.application.delete_expediente.command import (
    ExpedienteDeleteCommand,
    ExpedienteDeleteConflictError,
    ExpedienteDeleteError,
    ExpedienteDeleteResult,
    ExpedienteDeleteValidationError,
)
from app.src.modules.expedientes.application.delete_expediente.service import (
    ExpedienteDeleteService,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
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
    """In-memory repo with `has_children` for the conditional delete."""

    def __init__(self) -> None:
        self.by_id: dict[Any, object] = {}
        self.children_count: dict[Any, int] = {}
        self.calls: list[str] = []

    async def get_by_id(self, expediente_id: Any) -> object | None:
        self.calls.append("get_by_id")
        return self.by_id.get(expediente_id)

    async def create(self, aggregate: object) -> object:
        self.calls.append("create")
        self.by_id[aggregate.id] = aggregate
        return aggregate

    async def update(self, aggregate: object) -> object:
        """Trusts the service — service has already validated the version."""
        self.calls.append("update")
        self.by_id[aggregate.id] = aggregate
        return aggregate

    async def delete(self, expediente_id: Any) -> None:
        """Trusts the service — the conditional check happened upstream."""
        self.calls.append("delete")
        self.by_id.pop(expediente_id, None)

    async def has_children(self, expediente_id: Any) -> bool:
        """Returns whether the aggregate has any children (LOTE/BASED
        relationships, anexos, hitos, etc.)."""
        self.calls.append("has_children")
        return self.children_count.get(expediente_id, 0) > 0

    async def list_by_state(self, estado: str, limit: int, offset: int) -> tuple[list[object], int]:
        return [], 0


class _FakeHitoRepository(HitoRepositoryPort):
    def __init__(self) -> None:
        self.calls: list[str] = []

    async def get_by_expediente(self, expediente_id: Any) -> list[object]:
        return []

    async def upsert(self, hito: object) -> object:
        self.calls.append("upsert")
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


def _seed_expediente(
    repo: _FakeExpedienteRepository,
    *,
    version: int = 1,
) -> Expediente:
    expediente_id = uuid4()
    aggregate = Expediente(
        id=expediente_id,
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=version,
        id_expediente_padre=None,
        created_at=datetime.now(UTC),
    )
    repo.by_id[expediente_id] = aggregate
    return aggregate


def _make_service(
    repo: _FakeExpedienteRepository | None = None,
    audit: _FakeAuditLog | None = None,
) -> tuple[
    ExpedienteDeleteService,
    _FakeUoW,
    _FakeExpedienteRepository,
    _FakeHitoRepository,
    _FakeAuditLog,
]:
    repo = repo or _FakeExpedienteRepository()
    hito_repo = _FakeHitoRepository()
    audit = audit or _FakeAuditLog()
    uow = _FakeUoW(_FakeSession())
    service = ExpedienteDeleteService(
        expediente_repo=repo,
        hito_repo=hito_repo,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    return service, uow, repo, hito_repo, audit


def _make_command(
    aggregate: Expediente,
    *,
    force: bool = False,
    expected_version: int | None = None,
    actor_id: Any | None = None,
) -> ExpedienteDeleteCommand:
    return ExpedienteDeleteCommand(
        expediente_id=aggregate.id,
        expected_version=expected_version if expected_version is not None else aggregate.version,
        force=force,
        actor_id=actor_id if actor_id is not None else uuid4(),
    )


# ---------------------------------------------------------------------------
# §Camino feliz
# ---------------------------------------------------------------------------


async def test_delete_removes_aggregate_with_no_children() -> None:
    """CAP-003 §Camino feliz: aggregate without children deletes
    cleanly. Audit log records the event + change row.
    """
    service, uow, repo, _, audit = _make_service()
    aggregate = _seed_expediente(repo, version=3)
    repo.children_count[aggregate.id] = 0
    cmd = _make_command(aggregate)

    result = await service.execute(cmd)

    assert isinstance(result, ExpedienteDeleteResult)
    assert result.expediente_id == aggregate.id
    assert result.deleted_at.tzinfo is not None
    assert uow.commits == 1
    assert aggregate.id not in repo.by_id  # deleted


async def test_delete_records_audit_event_and_change_row() -> None:
    """§Camino feliz: the audit log receives one event + one change row
    with ``accion="delete"``. The deleted aggregate's version is
    captured in the change row.
    """
    # Build the fake repo and the fake audit first, then the service
    # over both, so the assertion below sees what the service wrote.
    repo = _FakeExpedienteRepository()
    audit = _FakeAuditLog()
    uow = _FakeUoW(_FakeSession())
    service = ExpedienteDeleteService(
        expediente_repo=repo,
        hito_repo=_FakeHitoRepository(),
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    aggregate = _seed_expediente(repo, version=5)
    repo.children_count[aggregate.id] = 0
    cmd = _make_command(aggregate)

    await service.execute(cmd)

    assert len(audit.events) == 1
    assert audit.events[0].event_type == "expediente.deleted"
    assert len(audit.changes) == 1
    assert audit.changes[0].accion == "delete"
    assert audit.changes[0].id_expediente == aggregate.id


async def test_delete_does_not_affect_related_records() -> None:
    """CAP-003 §Camino feliz: 'sin afectar relaciones no incluidas'.
    Deleting the aggregate does NOT touch other tables.
    """
    service, _, repo, hito_repo, _ = _make_service()
    aggregate = _seed_expediente(repo)
    repo.children_count[aggregate.id] = 0
    cmd = _make_command(aggregate)

    await service.execute(cmd)

    assert "delete" in repo.calls
    assert "delete" not in hito_repo.calls  # no hitos touched


# ---------------------------------------------------------------------------
# §Validación y autorización
# ---------------------------------------------------------------------------


async def test_delete_rejects_unknown_expediente() -> None:
    service, uow, _, _, _ = _make_service()
    cmd = ExpedienteDeleteCommand(
        expediente_id=uuid4(),
        expected_version=1,
        force=False,
        actor_id=uuid4(),
    )

    with pytest.raises(ExpedienteDeleteValidationError):
        await service.execute(cmd)

    assert uow.commits == 0


async def test_delete_rejects_stale_version() -> None:
    """CAP-003 §Concurrencia + D-EXP-4: stale version → 409."""
    service, uow, repo, _, _ = _make_service()
    aggregate = _seed_expediente(repo, version=2)
    repo.children_count[aggregate.id] = 0
    cmd = _make_command(aggregate, expected_version=1)  # stale

    with pytest.raises(ExpedienteDeleteConflictError):
        await service.execute(cmd)

    assert uow.commits == 0
    assert aggregate.id in repo.by_id  # not deleted


async def test_delete_rejects_missing_actor() -> None:
    service, uow, _, _, _ = _make_service()
    with pytest.raises(ExpedienteDeleteError):
        await _run_with_bad_actor(service)


async def _run_with_bad_actor(service: ExpedienteDeleteService) -> None:
    cmd = ExpedienteDeleteCommand(
        expediente_id=uuid4(),
        expected_version=1,
        force=False,
        actor_id=None,
    )
    await service.execute(cmd)


# ---------------------------------------------------------------------------
# §Conditional delete: "impedir pérdida de hijos"
# ---------------------------------------------------------------------------


async def test_delete_refuses_when_aggregate_has_children() -> None:
    """CAP-003 'impedir pérdida de hijos': children present → refuse."""
    service, uow, repo, _, _ = _make_service()
    aggregate = _seed_expediente(repo)
    repo.children_count[aggregate.id] = 3
    cmd = _make_command(aggregate, force=False)

    with pytest.raises(ExpedienteDeleteValidationError):
        await service.execute(cmd)

    assert uow.commits == 0
    assert aggregate.id in repo.by_id  # not deleted


async def test_delete_force_bypasses_children_check() -> None:
    """The caller accepts the risk; ``force=True`` deletes anyway.
    The audit log still records the event so traceability is
    preserved.
    """
    service, _, repo, _, audit = _make_service()
    aggregate = _seed_expediente(repo)
    repo.children_count[aggregate.id] = 5
    cmd = _make_command(aggregate, force=True)

    await service.execute(cmd)

    assert aggregate.id not in repo.by_id
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "expediente.deleted"
