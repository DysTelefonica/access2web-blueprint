"""Strict TDD for EXP-CAP-015..019 catalogs (Q01)."""

from __future__ import annotations

from dataclasses import dataclass, field
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.catalogs.command import (
    CatalogQuery,
    CatalogsAuthorizationError,
    CatalogsError,
    CatalogsValidationError,
    CpvQuery,
)
from app.src.modules.expedientes.application.catalogs.service import CatalogsService
from app.src.modules.expedientes.domain.cpv.code import CPVCode


@dataclass
class _FakeCatalogRow:
    id: int
    nombre: str
    descripcion: str = ""


@dataclass
class _CatalogRepo:
    by_id: dict[int, _FakeCatalogRow] = field(default_factory=dict)
    all_rows: list[_FakeCatalogRow] = field(default_factory=list)
    fail: bool = False

    async def list_all(self) -> list[object]:
        if self.fail:
            raise RuntimeError("database unavailable")
        return list(self.all_rows)

    async def get_by_id(self, catalog_id: int) -> object | None:
        if self.fail:
            raise RuntimeError("database unavailable")
        return self.by_id.get(catalog_id)

    async def search(self, query: str, limit: int = 20) -> list[object]:
        if self.fail:
            raise RuntimeError("database unavailable")
        lowered = query.lower()
        return [r for r in self.all_rows if lowered in r.nombre.lower()][:limit]


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


def _make_service(
    *,
    catalog_type: str = "comerciales",
    permission_required: str = "catalog.read",
) -> tuple[CatalogsService, _CatalogRepo, _Audit]:
    rows = [
        _FakeCatalogRow(id=1, nombre="Comercial Uno", descripcion="d1"),
        _FakeCatalogRow(id=2, nombre="Comercial Dos", descripcion="d2"),
    ]
    catalog = _CatalogRepo(
        by_id={row.id: row for row in rows},
        all_rows=rows,
    )
    audit = _Audit()
    service = CatalogsService(
        catalog_type=catalog_type,
        permission_required=permission_required,
        catalog_repo=catalog,
        audit_log=audit,
        permissions=set(),
    )
    return service, catalog, audit


def _actor_with(perms: set[str]) -> UUID:
    _ = perms
    return uuid4()


async def test_list_catalog_returns_rows_when_actor_has_permission() -> None:
    service, catalog, _ = _make_service()
    service._permissions = {"catalog.read"}  # bypass deny-by-default for the test
    query = CatalogQuery(actor_id=uuid4(), text=None, limit=20)

    rows = await service.list(query)

    assert [r.nombre for r in rows] == ["Comercial Uno", "Comercial Dos"]


async def test_list_catalog_rejects_actor_without_permission() -> None:
    service, _, _ = _make_service()

    with pytest.raises(CatalogsAuthorizationError, match="permission"):
        await service.list(CatalogQuery(actor_id=uuid4(), text=None, limit=20))


async def test_search_catalog_filters_by_text() -> None:
    service, catalog, _ = _make_service(catalog_type="ejercitos")
    service._permissions = {"catalog.read"}
    query = CatalogQuery(actor_id=uuid4(), text="dos", limit=20)

    rows = await service.list(query)

    assert [r.nombre for r in rows] == ["Comercial Dos"]


async def test_get_by_id_returns_entry_when_present() -> None:
    service, catalog, _ = _make_service()
    service._permissions = {"catalog.read"}

    row = await service.get(catalog_id=1, actor_id=uuid4())

    assert row.id == 1
    assert row.nombre == "Comercial Uno"


async def test_get_by_id_raises_when_missing() -> None:
    service, _, _ = _make_service()
    service._permissions = {"catalog.read"}

    with pytest.raises(CatalogsValidationError, match="not found"):
        await service.get(catalog_id=999, actor_id=uuid4())


async def test_dependency_failure_rolls_back_and_raises() -> None:
    service, catalog, _ = _make_service()
    catalog.fail = True
    service._permissions = {"catalog.read"}

    with pytest.raises(CatalogsError, match="catalogs failed"):
        await service.list(CatalogQuery(actor_id=uuid4(), text=None, limit=20))


async def test_cpv_decision_accepts_valid_catalog_entry() -> None:
    catalog = _CatalogRepo(
        by_id={3113100: _FakeCatalogRow(id=3113100, nombre="03113100-7", descripcion="Aceites")},
        all_rows=[],
    )
    service = CatalogsService(
        catalog_type="cpv",
        permission_required="catalog.read",
        catalog_repo=catalog,
        audit_log=_Audit(),
        permissions={"catalog.read"},
    )
    query = CpvQuery(
        actor_id=uuid4(),
        raw_code=CPVCode("03113100-7"),
    )

    decision = await service.cpv_decision(query)

    assert decision.accepted is True
    assert decision.row is not None
    assert decision.decision == "aceptado"


async def test_cpv_decision_rejects_unknown_code() -> None:
    catalog = _CatalogRepo(all_rows=[])
    service = CatalogsService(
        catalog_type="cpv",
        permission_required="catalog.read",
        catalog_repo=catalog,
        audit_log=_Audit(),
        permissions={"catalog.read"},
    )
    query = CpvQuery(
        actor_id=uuid4(),
        raw_code=CPVCode("99999999"),
    )

    decision = await service.cpv_decision(query)

    assert decision.accepted is False
    assert decision.decision == "rechazado"
    assert "99999999" in (decision.reason or "")


async def test_cpv_decision_without_permission_raises() -> None:
    catalog = _CatalogRepo(
        by_id={3113100: _FakeCatalogRow(id=3113100, nombre="03113100-7", descripcion="Aceites")},
    )
    service = CatalogsService(
        catalog_type="cpv",
        permission_required="catalog.read",
        catalog_repo=catalog,
        audit_log=_Audit(),
        permissions=set(),
    )
    query = CpvQuery(
        actor_id=uuid4(),
        raw_code=CPVCode("03113100-7"),
    )

    with pytest.raises(CatalogsAuthorizationError, match="permission"):
        await service.cpv_decision(query)


async def test_cpv_decision_audits_accepted_outcome() -> None:
    catalog = _CatalogRepo(
        by_id={3113100: _FakeCatalogRow(id=3113100, nombre="03113100-7", descripcion="Aceites")},
    )
    audit = _Audit()
    service = CatalogsService(
        catalog_type="cpv",
        permission_required="catalog.read",
        catalog_repo=catalog,
        audit_log=audit,
        permissions={"catalog.read"},
    )
    query = CpvQuery(
        actor_id=uuid4(),
        raw_code=CPVCode("03113100-7"),
    )

    await service.cpv_decision(query)

    assert len(audit.events) == 1
    assert audit.events[0].event_type == "cpv.decision"
    assert audit.events[0].payload["accepted"] is True
