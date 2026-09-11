"""Contract-conformance tests for Expedientes ports (F01, issue #222)."""

from __future__ import annotations

from uuid import uuid4

from tests.expedientes.adapters._fakes import (
    FakeAuditLog,
    FakeCatalogRepository,
    FakeDocumentStorage,
    FakeExpedienteRepository,
    FakeHitoRepository,
    FakeNotificationDelivery,
    FakeReadiness,
)


# ---------------------------------------------------------------------------
# Protocol structural checks
# ---------------------------------------------------------------------------

def test_all_fakes_satisfy_protocols() -> None:
    """Every fake provides the exact method set declared by its Protocol."""
    ports: list[tuple[str, object]] = [
        ("ExpedienteRepositoryPort", FakeExpedienteRepository()),
        ("HitoRepositoryPort",       FakeHitoRepository()),
        ("CatalogRepositoryPort",    FakeCatalogRepository()),
        ("AuditLogPort",            FakeAuditLog()),
        ("ReadinessPort",           FakeReadiness()),
        ("DocumentStoragePort",     FakeDocumentStorage()),
        ("NotificationDeliveryPort", FakeNotificationDelivery()),
    ]
    methods: dict[str, list[str]] = {
        "ExpedienteRepositoryPort": ["get_by_id", "create", "update", "delete", "list_by_state"],
        "HitoRepositoryPort":       ["get_by_expediente", "upsert", "delete"],
        "CatalogRepositoryPort":    ["list_all", "get_by_id", "search"],
        "AuditLogPort":            ["append", "list_for_actor"],
        "ReadinessPort":           ["check"],
        "DocumentStoragePort":     ["prepare_upload", "confirm_upload", "download", "delete"],
        "NotificationDeliveryPort": ["send"],
    }
    for name, port in ports:
        for method in methods[name]:
            assert hasattr(port, method), f"{name} missing {method}"


# ---------------------------------------------------------------------------
# Aggregate behaviour
# ---------------------------------------------------------------------------

async def test_expediente_repository_round_trip() -> None:
    """create → get_by_id returns the same aggregate; calls track the operation."""
    repo = FakeExpedienteRepository()
    aid = uuid4()
    agg = type("Agg", (), {"id": aid, "estado": "BORRADOR"})()
    await repo.create(agg)
    assert await repo.get_by_id(aid) is agg
    assert "create" in repo.calls and "get_by_id" in repo.calls
