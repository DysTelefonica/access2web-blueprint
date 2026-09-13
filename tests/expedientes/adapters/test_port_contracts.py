"""Contract-conformance tests for Expedientes ports (F01, issue #222).

Each test exercises one port through its Fake adapter against an actual domain
aggregate. The contract being verified is the *observable behaviour* declared
by the Protocol — not the duck-typed shape (already enforced by mypy on
``app/src/modules/expedientes/ports/``).
"""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import uuid4

from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo
from app.src.modules.expedientes.domain.hito import Hito
from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent
from app.src.modules.expedientes.ports.readiness import ReadinessResult
from tests.expedientes.adapters._fakes import (
    FakeAuditLog,
    FakeCatalogRepository,
    FakeDocumentStorage,
    FakeExpedienteRepository,
    FakeHitoRepository,
    FakeNotificationDelivery,
    FakeReadiness,
)


def _make_expediente() -> Expediente:
    """Real domain aggregate — AM type has no id_expediente_padre constraint."""
    return Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
        created_at=datetime(2026, 9, 11, tzinfo=UTC),
    )


# ---------------------------------------------------------------------------
# ExpedienteRepositoryPort (CAP-001..CAP-007, D-EXP-1, D-EXP-4)
# ---------------------------------------------------------------------------


async def test_expediente_repository_create_then_get_returns_same_aggregate() -> None:
    repo = FakeExpedienteRepository()
    exp = _make_expediente()
    created = await repo.create(exp)
    fetched = await repo.get_by_id(exp.id)

    assert fetched is exp
    assert created is exp
    assert repo.calls == ["create", "get_by_id"]


async def test_expediente_repository_list_by_state_filters_correctly() -> None:
    repo = FakeExpedienteRepository()
    borrador = _make_expediente()
    otro = Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.ADJUDICADO,
        version=1,
    )
    await repo.create(borrador)
    await repo.create(otro)

    rows, total = await repo.list_by_state("BORRADOR", limit=10, offset=0)
    assert total == 1
    assert rows == [borrador]


# ---------------------------------------------------------------------------
# HitoRepositoryPort (CAP-008, D-EXP-1)
# ---------------------------------------------------------------------------


async def test_hito_repository_upsert_attaches_to_expediente() -> None:
    from datetime import date

    repo = FakeHitoRepository()
    exp_id = uuid4()
    hito = Hito(
        id=uuid4(),
        id_expediente=exp_id,
        fecha_hito=date(2026, 12, 1),
        garantia_fecha_fin=None,
        estado=ExpedienteEstado.BORRADOR,
    )
    await repo.upsert(hito)
    fetched = await repo.get_by_expediente(exp_id)
    assert fetched == [hito]


# ---------------------------------------------------------------------------
# AuditLogPort (D-EXP-3, D-EXP-4, D55)
# ---------------------------------------------------------------------------


async def test_audit_log_append_and_list_for_actor_filters_by_actor() -> None:
    log = FakeAuditLog()
    actor = uuid4()
    other_actor = uuid4()
    e1 = ExpedienteAuditEvent(
        id=uuid4(),
        event_type="expediente.created",
        actor_id=actor,
        target_id=uuid4(),
        capacidad="EXP-CAP-001",
        created_at=datetime(2026, 9, 11, tzinfo=UTC),
    )
    e2 = ExpedienteAuditEvent(
        id=uuid4(),
        event_type="expediente.updated",
        actor_id=other_actor,
        target_id=uuid4(),
        capacidad="EXP-CAP-002",
        created_at=datetime(2026, 9, 11, tzinfo=UTC),
    )
    await log.append(e1)
    await log.append(e2)

    since = datetime(2026, 1, 1, tzinfo=UTC)
    for_actor = await log.list_for_actor(actor, since)
    assert for_actor == [e1]


# ---------------------------------------------------------------------------
# CatalogRepositoryPort (CAP-015..CAP-024)
# ---------------------------------------------------------------------------


async def test_catalog_repository_get_by_id_returns_stored_value() -> None:
    repo = FakeCatalogRepository()
    cat = type("Cat", (), {"descripcion": "Licitación abierta"})()
    repo.entries[42] = cat
    assert await repo.get_by_id(42) is cat
    assert await repo.get_by_id(99) is None


async def test_catalog_repository_search_filters_by_substring() -> None:
    repo = FakeCatalogRepository()
    repo.entries[1] = type("Cat", (), {"descripcion": "Licitación abierta"})()
    repo.entries[2] = type("Cat", (), {"descripcion": "Concurso público"})()
    results = await repo.search("licit")
    assert len(results) == 1


# ---------------------------------------------------------------------------
# ReadinessPort (CAP-047, D-EXP-3)
# ---------------------------------------------------------------------------


async def test_readiness_port_returns_result_with_checks() -> None:
    port = FakeReadiness()
    port._result = ReadinessResult(ready=True, checks=[])
    result = await port.check()
    assert isinstance(result, ReadinessResult)
    assert result.ready is True


# ---------------------------------------------------------------------------
# DocumentStoragePort (CAP-010, D-EXP-6)
# ---------------------------------------------------------------------------


async def test_document_storage_round_trip_returns_stored_bytes() -> None:
    storage = FakeDocumentStorage()
    exp_id = uuid4()
    ref = await storage.prepare_upload(exp_id, "doc.pdf", "application/pdf", 1024)
    stored = await storage.confirm_upload(ref)
    assert stored.storage_ref == ref
    assert stored.size_bytes == 1024
    content = await storage.download(ref)
    assert content == b""


# ---------------------------------------------------------------------------
# NotificationDeliveryPort (D-EXP-5, DA-10)
# ---------------------------------------------------------------------------


async def test_notification_delivery_records_message() -> None:
    notifier = FakeNotificationDelivery()
    await notifier.send("user@example.com", "Hola", "Cuerpo del mensaje")
    assert notifier.sent == [("user@example.com", "Hola", "Cuerpo del mensaje")]
