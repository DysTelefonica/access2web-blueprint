"""Strict TDD for EXP-CAP-025..029 query services (Q03)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.queries.command import (
    BandejaQuery,
    ExportacionExcelCommand,
    QueriesAuthorizationError,
    QueriesError,
    TareasCalculadasQuery,
    TareasResult,
)
from app.src.modules.expedientes.application.queries.service import QueriesService
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


@dataclass
class _FakeRow:
    id: UUID
    codigo: str
    estado: str
    tipo: str


@dataclass
class _ExpedienteRepo:
    all_rows: list[_FakeRow] = field(default_factory=list)
    fail: bool = False
    last_search_kwargs: dict[str, Any] = field(default_factory=dict)

    async def search(
        self,
        *,
        estado: str | None,
        codigo: str | None,
        responsable_id: object | None,
        juridica_id: object | None,
        suministrador_id: object | None,
        limit: int,
        offset: int,
    ) -> tuple[list[object], int]:
        if self.fail:
            raise RuntimeError("database unavailable")
        self.last_search_kwargs = {
            "estado": estado,
            "codigo": codigo,
            "responsable_id": responsable_id,
            "juridica_id": juridica_id,
            "suministrador_id": suministrador_id,
            "limit": limit,
            "offset": offset,
        }
        rows = list(self.all_rows)
        if estado is not None:
            rows = [r for r in rows if r.estado == estado]
        if codigo is not None:
            rows = [r for r in rows if codigo.lower() in r.codigo.lower()]
        if responsable_id is not None:
            rows = [r for r in rows if getattr(r, "responsable_id", None) == responsable_id]
        return rows[offset : offset + limit], len(rows)


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


def _make_service() -> tuple[QueriesService, _ExpedienteRepo, _Audit]:
    expediente = Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
    )
    repo = _ExpedienteRepo(
        all_rows=[
            _FakeRow(
                id=expediente.id,
                codigo="EXP-2026/01",
                estado=ExpedienteEstado.BORRADOR.value,
                tipo=ExpedienteTipo.AM.value,
            )
        ]
    )
    audit = _Audit()
    service = QueriesService(
        expediente_repo=repo,
        audit_log=audit,
        permissions=set(),
    )
    return service, repo, audit


async def test_bandeja_returns_page_with_total_count() -> None:
    service, repo, _ = _make_service()
    service.grant("bandeja.read")
    query = BandejaQuery(actor_id=uuid4(), estado=None, codigo=None, limit=20, offset=0)

    page = await service.bandeja(query)

    assert page.total == 1
    assert page.limit == 20
    assert page.offset == 0
    assert page.has_more is False


async def test_bandeja_forwards_filters_to_repository() -> None:
    service, repo, _ = _make_service()
    service.grant("bandeja.read")
    code = "EXP-2026/01"
    actor = uuid4()
    query = BandejaQuery(
        actor_id=actor,
        estado=ExpedienteEstado.BORRADOR.value,
        codigo=code,
        limit=10,
        offset=5,
    )

    await service.bandeja(query)

    assert repo.last_search_kwargs == {
        "estado": ExpedienteEstado.BORRADOR.value,
        "codigo": code,
        "responsable_id": None,
        "juridica_id": None,
        "suministrador_id": None,
        "limit": 10,
        "offset": 5,
    }


async def test_bandeja_rejects_actor_without_permission() -> None:
    service, _, _ = _make_service()
    query = BandejaQuery(actor_id=uuid4(), estado=None, codigo=None, limit=20, offset=0)

    with pytest.raises(QueriesAuthorizationError, match="bandeja.read"):
        await service.bandeja(query)


async def test_bandeja_dependency_failure_rolls_back() -> None:
    service, repo, audit = _make_service()
    repo.fail = True
    service.grant("bandeja.read")

    with pytest.raises(QueriesError, match="queries failed"):
        await service.bandeja(
            BandejaQuery(actor_id=uuid4(), estado=None, codigo=None, limit=20, offset=0)
        )

    assert not audit.events


async def test_exportacion_excel_produces_deterministic_payload() -> None:
    service, _, audit = _make_service()
    service.grant("export.excel")
    cmd = ExportacionExcelCommand(actor_id=uuid4(), filtros={}, formato="xlsx")

    job = await service.exportar_excel(cmd)

    assert job.formato == "xlsx"
    assert job.filtros == {}
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "export.excel.requested"


async def test_exportacion_excel_rejects_missing_permission() -> None:
    service, _, _ = _make_service()
    cmd = ExportacionExcelCommand(actor_id=uuid4(), filtros={}, formato="xlsx")

    with pytest.raises(QueriesAuthorizationError):
        await service.exportar_excel(cmd)


async def test_tareas_calculadas_six_buckets_zero_when_empty() -> None:
    service, _, _ = _make_service()
    service.grant("tareas.read")

    result = await service.tareas_calculadas(
        TareasCalculadasQuery(actor_id=uuid4(), now=__import__("datetime").datetime.now())
    )

    assert isinstance(result, TareasResult)
    assert sum(result.contadores.values()) == 0
    assert set(result.contadores) == {
        "desconocido",
        "recepcion_hito",
        "adjudicado_sin_contrato",
        "tsol_sin_s4h",
        "oferta_prolongada",
        "otro",
    }


async def test_tareas_calculadas_categorize_borrador_as_desconocido() -> None:
    service, _, _ = _make_service()
    service.grant("tareas.read")
    expediente = Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
    )
    service.bind_expedientes(lambda: [expediente])

    result = await service.tareas_calculadas(
        TareasCalculadasQuery(actor_id=uuid4(), now=__import__("datetime").datetime.now())
    )

    assert result.contadores["desconocido"] == 1
    assert sum(result.contadores.values()) == 1
