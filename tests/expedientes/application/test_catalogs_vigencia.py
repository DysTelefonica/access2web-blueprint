"""Strict TDD for Q02 catalog group B vigencia (CAP-020..024)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from uuid import uuid4

import pytest

from app.src.modules.expedientes.application.catalogs.command import (
    CatalogQuery,
    CatalogsAuthorizationError,
    CatalogsError,
)
from app.src.modules.expedientes.application.catalogs.service import CatalogsService


@dataclass
class _CatalogRow:
    id: int
    nombre: str
    descripcion: str = ""


@dataclass
class _VigenteCatalogRepo:
    by_id: dict[int, _CatalogRow] = field(default_factory=dict)
    all_rows: list[_CatalogRow] = field(default_factory=list)
    vigentes_cache: list[_CatalogRow] = field(default_factory=list)
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

    async def list_vigentes_en(self, fecha: date) -> list[object]:
        if self.fail:
            raise RuntimeError("database unavailable")
        return list(self.vigentes_cache)

    async def vigentes_search(self, query: str, fecha: date, limit: int = 20) -> list[object]:
        if self.fail:
            raise RuntimeError("database unavailable")
        lowered = query.lower()
        return [r for r in self.vigentes_cache if lowered in r.nombre.lower()][:limit]


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
    catalog_type: str = "pecal",
    permission: str = "catalog.read",
    vigentes_cache: list[_CatalogRow] | None = None,
) -> tuple[CatalogsService, _VigenteCatalogRepo, _Audit]:
    rows = [
        _CatalogRow(id=1, nombre="PECAL UNO", descripcion="d1"),
        _CatalogRow(id=2, nombre="PECAL DOS", descripcion="d2"),
    ]
    catalog = _VigenteCatalogRepo(
        by_id={r.id: r for r in rows},
        all_rows=rows,
        vigentes_cache=vigentes_cache if vigentes_cache is not None else list(rows),
    )
    audit = _Audit()
    service = CatalogsService(
        catalog_type=catalog_type,
        permission_required=permission,
        catalog_repo=catalog,
        audit_log=audit,
        permissions=set(),
    )
    return service, catalog, audit


async def test_pecal_list_with_fecha_uses_vigentes_path() -> None:
    visible = [_CatalogRow(id=2, nombre="PECAL DOS", descripcion="d2")]
    service, _, _ = _make_service(vigentes_cache=visible)
    service.grant("catalog.read")

    rows = await service.list(CatalogQuery(actor_id=uuid4(), text=None, fecha=date(2026, 6, 1)))

    assert [r.id for r in rows] == [2]


async def test_pecal_search_with_fecha_uses_vigentes_search() -> None:
    visible = [
        _CatalogRow(id=1, nombre="PECAL UNO", descripcion="d1"),
        _CatalogRow(id=2, nombre="PECAL DOS", descripcion="d2"),
    ]
    service, _, _ = _make_service(vigentes_cache=visible)
    service.grant("catalog.read")

    rows = await service.list(CatalogQuery(actor_id=uuid4(), text="uno", fecha=date(2026, 6, 1)))

    assert [r.id for r in rows] == [1]


async def test_racs_list_audits_with_racs_capability() -> None:
    service, _, audit = _make_service(catalog_type="racs", permission="catalog.read")
    service.grant("catalog.read")

    await service.list(CatalogQuery(actor_id=uuid4(), text=None, fecha=date(2026, 6, 1)))

    assert len(audit.events) == 1
    assert audit.events[0].capacidad == "EXP-CAP-021"


async def test_unknown_catalog_type_is_rejected() -> None:
    service, _, _ = _make_service()

    with pytest.raises(ValueError, match="unknown catalog_type"):
        CatalogsService(
            catalog_type="unknown",
            permission_required="catalog.read",
            catalog_repo=_VigenteCatalogRepo(),
            audit_log=_Audit(),
            permissions=set(),
        )


async def test_vigencia_dependency_failure_rolls_back() -> None:
    service, catalog, audit = _make_service()
    catalog.fail = True
    service.grant("catalog.read")

    with pytest.raises(CatalogsError, match="catalogs failed"):
        await service.list(CatalogQuery(actor_id=uuid4(), text=None, fecha=date(2026, 6, 1)))

    assert not audit.events


async def test_vigencia_query_without_permission_raises() -> None:
    service, _, _ = _make_service(catalog_type="oficinas_programa", permission="catalog.read")

    with pytest.raises(CatalogsAuthorizationError, match="oficinas_programa"):
        await service.list(CatalogQuery(actor_id=uuid4(), text=None, fecha=date(2026, 6, 1)))


async def test_vigencia_query_without_fecha_falls_back_to_list_all() -> None:
    """Without ``fecha`` the service uses ``list_all`` (the
    pre-vigencia path). The adapter decides what to do; the
    service just keeps the read path auditable.
    """
    service, _, audit = _make_service()
    service.grant("catalog.read")

    rows = await service.list(CatalogQuery(actor_id=uuid4(), text=None, fecha=None))

    assert [r.id for r in rows] == [1, 2]
    assert len(audit.events) == 1
