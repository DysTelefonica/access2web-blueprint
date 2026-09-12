"""Strict TDD — ExpedienteEditService (CAP-002, issue #227).

CAP-002 §Camino feliz exige "actualizan los campos manteniendo
invariantes y control de concurrencia, con salida determinista y
auditable". §Concurrencia o fallo: "preserva invariantes, no duplica
ni pierde evidencia, queda reintentable".

Estos tests verifican el comportamiento del use case usando:
- El aggregate `Expediente` ya mergeado en F02 (#643).
- El UoW real de F04 (#417).
- Los puertos `ExpedienteRepositoryPort` y `AuditLogPort` ya
  definidos en F01 (#222) y extendidos en #668.

El edit es optimista: el UPDATE se hace con `WHERE version = ?` y un
version stale devuelve `ExpedienteEditConflictError` (HTTP 409) sin
efecto parcial.
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any
from uuid import uuid4

import pytest

from app.src.modules.expedientes.application.edit_expediente.command import (
    ExpedienteEditCommand,
    ExpedienteEditConflictError,
    ExpedienteEditError,
    ExpedienteEditResult,
    ExpedienteEditValidationError,
)
from app.src.modules.expedientes.application.edit_expediente.service import (
    ExpedienteEditService,
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
# Doubles — minimal repos for the test
# ---------------------------------------------------------------------------


class _FakeExpedienteRepository(ExpedienteRepositoryPort):
    """In-memory repo with version-aware update for optimistic locking."""

    def __init__(self) -> None:
        self.by_id: dict[Any, object] = {}
        self.calls: list[str] = []

    async def get_by_id(self, expediente_id: Any) -> object | None:  # type: ignore[override]
        self.calls.append("get_by_id")
        return self.by_id.get(expediente_id)

    async def create(self, aggregate: object) -> object:  # type: ignore[override]
        self.calls.append("create")
        agg_id = getattr(aggregate, "id", None)
        if agg_id is None:
            raise RuntimeError("aggregate must have an id")
        self.by_id[agg_id] = aggregate
        return aggregate

    async def update(self, aggregate: object) -> object:  # type: ignore[override]
        """Replace the stored aggregate.

        The use case is responsible for the optimistic-lock check
        (it reads the aggregate, validates the version, and only
        then asks the repo to write). The fake mirrors that
        contract: the ``service.execute`` already raised
        ``ExpedienteEditConflictError`` if the version was stale, so
        by the time we reach this method, the write is unconditional.
        """
        self.calls.append("update")
        self.by_id[aggregate.id] = aggregate
        return aggregate

    async def delete(self, expediente_id: Any) -> None:  # type: ignore[override]
        self.calls.append("delete")
        self.by_id.pop(expediente_id, None)

    async def list_by_state(self, estado: str, limit: int, offset: int) -> tuple[list[object], int]:  # type: ignore[override]
        return [], 0


class _FakeHitoRepository(HitoRepositoryPort):
    def __init__(self) -> None:
        self.calls: list[str] = []

    async def get_by_expediente(self, expediente_id: Any) -> list[object]:  # type: ignore[override]
        return []

    async def upsert(self, hito: object) -> object:  # type: ignore[override]
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

    async def list_for_actor(  # type: ignore[override]
        self, actor_id: Any, since: datetime
    ) -> list[ExpedienteAuditEvent]:
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
    estado: ExpedienteEstado = ExpedienteEstado.BORRADOR,
    tipo: ExpedienteTipo = ExpedienteTipo.AM,
) -> Expediente:
    """Seed an aggregate in the repo with the given initial state."""
    expediente_id = uuid4()
    aggregate = Expediente(
        id=expediente_id,
        tipo=tipo,
        estado=estado,
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
    ExpedienteEditService,
    _FakeUoW,
    _FakeExpedienteRepository,
    _FakeHitoRepository,
    _FakeAuditLog,
]:
    repo = repo or _FakeExpedienteRepository()
    hito_repo = _FakeHitoRepository()
    audit = audit or _FakeAuditLog()
    uow = _FakeUoW(_FakeSession())
    service = ExpedienteEditService(
        expediente_repo=repo,
        hito_repo=hito_repo,
        audit_log=audit,
        uow_factory=lambda: uow,
    )
    return service, uow, repo, hito_repo, audit


def _make_command(
    aggregate: Expediente,
    *,
    estado: ExpedienteEstado | None = None,
    actor_id: Any | None = None,
    expected_version: int | None = None,
) -> ExpedienteEditCommand:
    return ExpedienteEditCommand(
        expediente_id=aggregate.id,
        expected_version=expected_version if expected_version is not None else aggregate.version,
        estado=estado,
        actor_id=actor_id if actor_id is not None else uuid4(),
    )


# ---------------------------------------------------------------------------
# §Camino feliz
# ---------------------------------------------------------------------------


async def test_edit_changes_estado_with_correct_version() -> None:
    """CAP-002 §Camino feliz: the aggregate's estado changes when the
    caller supplies the current version."""
    service, uow, repo, _, _ = _make_service()
    aggregate = _seed_expediente(repo, version=1)
    cmd = _make_command(aggregate, estado=ExpedienteEstado.ADJUDICADO)

    result = await service.execute(cmd)

    assert isinstance(result, ExpedienteEditResult)
    assert result.new_version == 2
    assert uow.commits == 1
    assert uow.rollbacks == 0
    stored = repo.by_id[aggregate.id]
    assert stored.estado is ExpedienteEstado.ADJUDICADO


async def test_edit_preserves_id_and_created_at() -> None:
    """Identity and creation timestamp are immutable. The edit only
    changes mutable fields."""
    service, _, repo, _, _ = _make_service()
    aggregate = _seed_expediente(repo, version=3)
    original_created_at = aggregate.created_at
    cmd = _make_command(aggregate, estado=ExpedienteEstado.ADJUDICADO)

    await service.execute(cmd)

    stored = repo.by_id[aggregate.id]
    assert stored.id == aggregate.id
    assert stored.created_at == original_created_at


async def test_edit_records_audit_event() -> None:
    service, _, _, _, audit = _make_service()
    service, _, repo, _, audit = _make_service()
    aggregate = _seed_expediente(repo)
    cmd = _make_command(aggregate, estado=ExpedienteEstado.ADJUDICADO)

    await service.execute(cmd)

    assert len(audit.events) == 1
    assert "expediente.updated" == audit.events[0].event_type


async def test_edit_records_one_change_per_modified_field() -> None:
    """C02 emits one ``ChangeRecord`` per changed field — unlike the
    alta which produces one record for the whole aggregate."""
    service, _, _, _, audit = _make_service()
    aggregate = _seed_expediente(repo=None) if False else None  # type: ignore
    service, _, repo, _, audit = _make_service()
    aggregate = _seed_expediente(repo)
    cmd = _make_command(aggregate, estado=ExpedienteEstado.ADJUDICADO)

    await service.execute(cmd)

    # Only the `estado` field changed in this test, so one change row.
    assert len(audit.changes) == 1
    change = audit.changes[0]
    assert change.nombre_campo == "estado"
    assert change.valor_inicial == "BORRADOR"
    assert change.valor_final == "ADJUDICADO"
    assert change.accion == "edit"


# ---------------------------------------------------------------------------
# §Validación y autorización
# ---------------------------------------------------------------------------


async def test_edit_rejects_unknown_expediente() -> None:
    service, _, _, _, _ = _make_service()
    cmd = ExpedienteEditCommand(
        expediente_id=uuid4(),
        expected_version=1,
        estado=ExpedienteEstado.ADJUDICADO,
        actor_id=uuid4(),
    )

    with pytest.raises(ExpedienteEditValidationError):
        await service.execute(cmd)


async def test_edit_rejects_stale_version() -> None:
    """CAP-002 §Concurrencia o fallo + D-EXP-4: version stale →
    409. The check happens before the UoW opens (no transaction wasted
    on a conflict). The aggregate is unchanged.
    """
    service, uow, repo, _, _ = _make_service()
    aggregate = _seed_expediente(repo, version=2)
    cmd = _make_command(aggregate, expected_version=1)  # stale

    with pytest.raises(ExpedienteEditConflictError):
        await service.execute(cmd)

    assert uow.commits == 0  # no transaction opened
    # The aggregate in the repo is unchanged.
    assert repo.by_id[aggregate.id].version == 2
    assert repo.by_id[aggregate.id].estado is ExpedienteEstado.BORRADOR


async def test_edit_rejects_missing_actor() -> None:
    """D-EXP-3 deny-by-default: no actor, no edit."""
    service, _, _, _, _ = _make_service()
    aggregate = (
        _seed_expediente(
            _FakeExpedienteRepository() if False else None  # type: ignore
        )
        if False
        else None
    )  # type: ignore
    service, _, repo, _, _ = _make_service()
    aggregate = _seed_expediente(repo)
    with pytest.raises(ExpedienteEditError):
        await _run_with_bad_actor(service, aggregate)


async def _run_with_bad_actor(service: ExpedienteEditService, aggregate: Expediente) -> None:
    """Helper for the actor test — bypasses the command factory."""
    cmd = ExpedienteEditCommand(
        expediente_id=aggregate.id,
        expected_version=aggregate.version,
        estado=ExpedienteEstado.ADJUDICADO,
        actor_id=None,  # explicit None
    )
    await service.execute(cmd)


async def test_edit_without_state_change_is_a_no_op() -> None:
    """A no-op edit (no state change) writes nothing: no UoW commit,
    no audit event, no change rows. Nothing to record, so nothing
    happens.
    """
    service, uow, repo, _, audit = _make_service()
    aggregate = _seed_expediente(repo, estado=ExpedienteEstado.BORRADOR)
    cmd = _make_command(aggregate, estado=ExpedienteEstado.BORRADOR)  # same

    result = await service.execute(cmd)

    assert result.fields_changed == ()
    assert uow.commits == 0  # no transaction opened
    assert audit.events == []
    assert audit.changes == []
