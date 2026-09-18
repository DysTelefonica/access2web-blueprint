"""UAT end-to-end for CAP-015..024 catalogs (issue #263, U03)."""

from dataclasses import dataclass, field
from uuid import uuid4

import pytest

from app.src.modules.expedientes.application.catalogs.command import (
    CatalogQuery,
    CatalogsAuthorizationError,
    CatalogsValidationError,
    CpvQuery,
)
from app.src.modules.expedientes.application.catalogs.service import CatalogsService
from app.src.modules.expedientes.domain.cpv.code import CPVCode

CATALOG_TYPES = [
    "comerciales",
    "cpv",
    "ejercitos",
    "suministradores",
    "lugares_ejecucion",
    "pecal",
    "racs",
    "grados",
    "organos",
    "oficinas_programa",
]


@dataclass
class _FakeRow:
    id: int
    nombre: str
    descripcion: str = ""


@dataclass
class _AuditLog:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        return None


@dataclass
class _CatalogRepo:
    by_id: dict[int, _FakeRow] = field(default_factory=dict)
    all_rows: list[_FakeRow] = field(default_factory=list)
    vigentes: list[_FakeRow] = field(default_factory=list)

    async def list_all(self) -> list[object]:
        return list(self.all_rows)

    async def get_by_id(self, catalog_id: int) -> object | None:
        return self.by_id.get(catalog_id)

    async def search(self, query: str, limit: int = 20) -> list[object]:
        return [r for r in self.all_rows if query.lower() in r.nombre.lower()][:limit]

    async def vigentes_search(self, query: str, fecha: object, limit: int = 20) -> list[object]:
        return list(self.vigentes)[:limit]


def _service(
    audit: _AuditLog, *, catalog_type: str, permission: str, repo: _CatalogRepo
) -> CatalogsService:
    return CatalogsService(
        catalog_type=catalog_type,
        permission_required=permission,
        catalog_repo=repo,
        audit_log=audit,
        permissions={permission},
    )


def _types(audit: _AuditLog) -> list[str]:
    return [e.event_type for e in audit.events]  # type: ignore[attr-defined]


def _caps(audit: _AuditLog) -> list[str]:
    return [e.capacidad for e in audit.events]  # type: ignore[attr-defined]


_EXPECTED_CAPS_BY_TYPE = {
    "comerciales": "EXP-CAP-015",
    "cpv": "EXP-CAP-016",
    "ejercitos": "EXP-CAP-017",
    "suministradores": "EXP-CAP-018",
    "lugares_ejecucion": "EXP-CAP-019",
    "pecal": "EXP-CAP-020",
    "racs": "EXP-CAP-021",
    "grados": "EXP-CAP-022",
    "organos": "EXP-CAP-023",
    "oficinas_programa": "EXP-CAP-024",
}


async def test_list_across_ten_catalogs_emits_correct_capacidad_per_type() -> None:
    audit = _AuditLog()
    actor = uuid4()
    row = _FakeRow(id=1, nombre="placeholder")

    for cat_type in CATALOG_TYPES:
        repo = _CatalogRepo(all_rows=[row])
        svc = _service(audit, catalog_type=cat_type, permission=f"{cat_type}.read", repo=repo)
        await svc.list(CatalogQuery(actor_id=actor, text=None))

    assert _types(audit) == ["catalog.queried"] * len(CATALOG_TYPES)
    assert _caps(audit) == [_EXPECTED_CAPS_BY_TYPE[t] for t in CATALOG_TYPES]


async def test_cpv_decision_emits_cpv_event_for_cpv_catalog() -> None:
    audit = _AuditLog()
    actor = uuid4()
    svc = _service(
        audit,
        catalog_type="cpv",
        permission="cpv.read",
        repo=_CatalogRepo(by_id={12345678: _FakeRow(id=12345678, nombre="X")}),
    )

    decision = await svc.cpv_decision(CpvQuery(actor_id=actor, raw_code=CPVCode("12345678")))

    assert decision.accepted is True
    assert _types(audit) == ["cpv.decision"]
    assert _caps(audit) == ["EXP-CAP-016"]


async def test_cpv_decision_rejects_non_cpv_catalog_type() -> None:
    audit = _AuditLog()
    actor = uuid4()
    svc = _service(
        audit, catalog_type="comerciales", permission="comerciales.read", repo=_CatalogRepo()
    )
    with pytest.raises(CatalogsValidationError, match="cpv"):
        await svc.cpv_decision(CpvQuery(actor_id=actor, raw_code=CPVCode("12345678")))


async def test_missing_actor_raises_authorization() -> None:
    audit = _AuditLog()
    svc = _service(
        audit, catalog_type="comerciales", permission="comerciales.read", repo=_CatalogRepo()
    )
    with pytest.raises(CatalogsAuthorizationError, match="actor_id"):
        await svc.list(CatalogQuery(actor_id=None, text=None))
    assert audit.events == []
