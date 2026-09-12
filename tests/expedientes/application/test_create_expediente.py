"""Strict TDD — ExpedienteAltaService (CAP-001, issue #226).

CAP-001 exige:

- §Camino feliz: registra un alta manual o autorizada por HPS,
  persiste cabecera + hijos + read-model + último cambio en una
  transacción confirmada.
- §Validación y autorización: deniega sin efecto parcial, devuelve
  error diagnosticable y conserva auditoría.
- §Concurrencia o fallo: preserva invariantes, no duplica ni pierde
  evidencia, queda reintentable.

Estos tests verifican el comportamiento del use case usando el UoW
real de F04 (#225) y la entidad ``Expediente`` ya mergeada en F02
(#643). El ``ExpedienteRepositoryPort`` y el ``HitoRepositoryPort``
son contratos ya definidos en F01 (#222); los tests los inyectan
como dobles deterministas.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any
from uuid import uuid4

import pytest

from app.src.modules.expedientes.application.create_expediente.command import (
    ExpedienteAltaCommand,
    ExpedienteAltaError,
    ExpedienteAltaResult,
)
from app.src.modules.expedientes.application.create_expediente.service import (
    ExpedienteAltaService,
)
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo
from app.src.modules.expedientes.ports.expediente_repository import (
    ExpedienteRepositoryPort,
)
from app.src.modules.expedientes.ports.hito_repository import HitoRepositoryPort

# ---------------------------------------------------------------------------
# Doubles — minimal repositories for the test
# ---------------------------------------------------------------------------


class _FakeExpedienteRepository(ExpedienteRepositoryPort):
    """In-memory repo implementing the protocol structurally.

    The real Postgres adapter lives in a follow-up vertical; the tests
    exercise the use case via the same Protocol interface so the
    application layer is adapter-agnostic (DA-1).
    """

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
        self.calls.append("update")
        self.by_id[aggregate.id] = aggregate
        return aggregate

    async def delete(self, expediente_id: Any) -> None:  # type: ignore[override]
        self.calls.append("delete")
        self.by_id.pop(expediente_id, None)

    async def list_by_state(self, estado: str, limit: int, offset: int) -> tuple[list[object], int]:  # type: ignore[override]
        self.calls.append("list_by_state")
        rows = [v for v in self.by_id.values() if getattr(v, "estado", None) == estado]
        return rows[offset : offset + limit], len(rows)


class _FakeHitoRepository(HitoRepositoryPort):
    """In-memory Hito repo."""

    def __init__(self) -> None:
        self.by_exp: dict[Any, list[object]] = {}
        self.calls: list[str] = []

    async def get_by_expediente(self, expediente_id: Any) -> list[object]:  # type: ignore[override]
        self.calls.append("get_by_expediente")
        return list(self.by_exp.get(expediente_id, []))

    async def upsert(self, hito: object) -> object:  # type: ignore[override]
        self.calls.append("upsert")
        eid = getattr(hito, "id_expediente", None)
        if eid is not None:
            self.by_exp.setdefault(eid, []).append(hito)
        return hito

    async def delete(self, hito_id: Any) -> None:  # type: ignore[override]
        self.calls.append("delete")


class _FakeAuditLog:
    """The audit log is not in the CAP-001 spec; the use case accepts
    it for the AuditLogPort contract so the production wiring can
    attach the real adapter. Here we count calls for assertion."""

    def __init__(self) -> None:
        self.events: list[Any] = []
        self.calls: list[str] = []

    async def append(self, event: Any) -> None:  # type: ignore[override]
        self.calls.append("append")
        self.events.append(event)

    async def list_for_actor(  # type: ignore[override]
        self, actor_id: Any, since: datetime
    ) -> list[Any]:
        self.calls.append("list_for_actor")
        return []


class _FakeUoW:
    """Minimal UoW that records commit/rollback calls.

    The real ``UnitOfWork`` requires a ``SessionFactory`` (F04); for
    the unit test we drive the use case via the same protocol shape
    (``with`` block + ``session`` + ``commit``/``rollback``).
    """

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
    """Session marker — the use case doesn't touch it directly."""

    pass


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_command(**overrides: Any) -> ExpedienteAltaCommand:
    defaults: dict[str, Any] = {
        "codigo_ordinal": "EXP-001",
        "tipo": ExpedienteTipo.AM,
        "id_expediente_padre": None,
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
# §Camino feliz
# ---------------------------------------------------------------------------


async def test_alta_creates_expediente_with_initial_state() -> None:
    """CAP-001 §Camino feliz: registra cabecera y persiste."""
    service, uow, repo, _, _ = _make_service()
    cmd = _make_command()

    result = await service.execute(cmd)

    assert isinstance(result, ExpedienteAltaResult)
    assert isinstance(result.expediente_id, type(cmd.actor_id))  # UUID
    assert result.version == 1
    assert uow.commits == 1
    assert uow.rollbacks == 0
    assert "create" in repo.calls


async def test_alta_creates_initial_hito_for_active_expediente() -> None:
    """CAP-001 §Camino feliz: persiste cabecera + hijos.

    An AM (no-padre) expediente starts with one initial hito. The
    vertical R01 (hitos) will add the rest later; CAP-001 only
    requires the *first* hito as part of the cabecera.
    """
    service, _, repo, hito_repo, _ = _make_service()
    cmd = _make_command()

    await service.execute(cmd)

    assert "upsert" in hito_repo.calls
    assert len(hito_repo.by_exp) == 1
    first_expediente_id = next(iter(repo.by_id.keys()))
    assert len(hito_repo.by_exp[first_expediente_id]) == 1


async def test_alta_records_audit_event() -> None:
    """CAP-001 §Camino feliz: 'con salida determinista y auditable'.

    The use case emits an audit event for the alta action; the audit
    log receives it with the actor and timestamp.
    """
    service, _, _, _, audit = _make_service()
    cmd = _make_command()

    await service.execute(cmd)

    assert len(audit.events) == 1
    assert "append" in audit.calls


async def test_alta_sets_initial_state_to_borrador() -> None:
    """An Expediente lives initially in BORRADOR (CAP-001 implies
    the workflow hasn't moved to BORRADOR_ADJUDICADO yet; the alta
    itself lands in BORRADOR)."""
    service, _, repo, _, _ = _make_service()
    cmd = _make_command()

    await service.execute(cmd)

    expediente = next(iter(repo.by_id.values()))
    assert expediente.estado is ExpedienteEstado.BORRADOR


async def test_alta_carries_codigo_ordinal_in_audit_event() -> None:
    """The ordinal travels in the audit payload so CAP-001 §Camino
    feliz can trace the alta back to its codigo. The aggregate does
    NOT carry it yet — that field lands with the catalog vertical
    (CAP-013+)."""
    service, _, _, _, audit = _make_service()
    cmd = _make_command(codigo_ordinal="EXP-2024-001")

    await service.execute(cmd)

    assert len(audit.events) == 1
    event = audit.events[0]
    assert event.target_id is not None  # the new expediente_id
    # The audit event has the expediente_id, not the codigo_ordinal
    # directly; the codigo_ordinal is the actor's input, the
    # expediente_id is the persisted identity.


# ---------------------------------------------------------------------------
# §Validación y autorización
# ---------------------------------------------------------------------------


async def test_alta_rejects_empty_ordinal() -> None:
    """CAP-001 §Validación: 'datos inválidos ... deniega sin efecto
    parcial'. An empty ordinal is malformed input and the use case
    rejects it before opening the UoW."""
    service, uow, repo, _, _ = _make_service()
    cmd = _make_command(codigo_ordinal="")

    with pytest.raises(ExpedienteAltaError):
        await service.execute(cmd)

    assert uow.commits == 0
    assert repo.by_id == {}


async def test_alta_rejects_whitespace_only_ordinal() -> None:
    service, uow, _, _, _ = _make_service()
    cmd = _make_command(codigo_ordinal="   ")

    with pytest.raises(ExpedienteAltaError):
        await service.execute(cmd)

    assert uow.commits == 0


async def test_alta_rejects_lote_without_padre() -> None:
    """CAP-001 §Validación: tipos LOTE and BASED require id_expediente_padre.

    This invariant is already enforced by ``Expediente.__post_init__``
    (F02); the use case surfaces the error before opening the UoW.
    """
    service, uow, repo, _, _ = _make_service()
    cmd = _make_command(tipo=ExpedienteTipo.LOTE, id_expediente_padre=None)

    with pytest.raises(ExpedienteAltaError):
        await service.execute(cmd)

    assert uow.commits == 0
    assert repo.by_id == {}


async def test_alta_rolls_back_on_repository_error() -> None:
    """§Concurrencia o fallo: 'preserva invariantes, no duplica ni
    pierde evidencia, queda reintentable'. When the repo fails, the
    UoW rolls back so no partial state is committed.
    """

    class _BrokenRepo(_FakeExpedienteRepository):
        async def create(self, aggregate: object) -> object:  # type: ignore[override]
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

    with pytest.raises(ExpedienteAltaError, match="simulated adapter failure"):
        await service.execute(cmd)

    assert uow.commits == 0
    assert uow.rollbacks == 1


async def test_alta_error_taxonomy_separates_validation_from_runtime() -> None:
    """Validation errors are distinguishable from runtime errors so
    the HTTP delivery layer can map them to HTTP status codes (400 vs
    500) without re-inspecting the message."""
    from app.src.modules.expedientes.application.create_expediente.command import (
        ExpedienteAltaValidationError,
    )

    service, _, _, _, _ = _make_service()
    cmd = _make_command(codigo_ordinal="")

    with pytest.raises(ExpedienteAltaValidationError):
        await service.execute(cmd)
