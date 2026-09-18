"""UAT end-to-end for CAP-025..029 query-and-tasks (issue #277, U04)."""

from dataclasses import dataclass, field
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.queries.command import (
    BandejaQuery,
    ExportacionExcelCommand,
    TareasCalculadasQuery,
)
from app.src.modules.expedientes.application.queries.service import QueriesService
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


@dataclass
class _AuditLog:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        return None


class _Repo:
    def __init__(self, rows: list[Expediente]) -> None:
        self._rows = rows

    async def search(self, **_kw: Any) -> tuple[list[Expediente], int]:
        return (self._rows, len(self._rows))

    async def list_active(self, **_kw: Any) -> list[Expediente]:
        return self._rows

    async def get_by_id(self, eid: UUID) -> Expediente | None:
        return next((r for r in self._rows if r.id == eid), None)


def _stack(audit: _AuditLog, rows: list[Expediente]) -> QueriesService:
    return QueriesService(
        expediente_repo=_Repo(rows),
        audit_log=audit,
        permissions={"bandeja.read", "bandeja.technical", "export.excel", "tareas.read"},
    )


def _types(audit: _AuditLog) -> list[str]:
    return [e.event_type for e in audit.events]  # type: ignore[attr-defined]


def _make_exp() -> Expediente:
    return Expediente(
        id=uuid4(), tipo=ExpedienteTipo.AM, estado=ExpedienteEstado.BORRADOR, version=1
    )


async def test_full_chain_emits_three_audit_events_for_queries_and_tasks() -> None:
    audit, actor = _AuditLog(), uuid4()
    rows = [_make_exp(), _make_exp()]
    svc = _stack(audit, rows)

    await svc.bandeja(BandejaQuery(actor_id=actor, estado=None, codigo=None))
    await svc.tareas_calculadas(
        TareasCalculadasQuery(
            actor_id=actor, now=__import__("datetime").datetime.now(__import__("datetime").UTC)
        )
    )
    await svc.exportar_excel(ExportacionExcelCommand(actor_id=actor, filtros={}))

    assert _types(audit) == ["bandeja.queried", "tareas.calculated", "export.excel.requested"]


async def test_three_query_methods_share_one_audit_log_and_emit_distinct_event_ids() -> None:
    audit, actor = _AuditLog(), uuid4()
    svc = _stack(audit, [_make_exp()])

    await svc.bandeja(BandejaQuery(actor_id=actor, estado=None, codigo=None))
    await svc.busqueda_avanzada(BandejaQuery(actor_id=actor, estado=None, codigo=None))
    await svc.tareas_calculadas(
        TareasCalculadasQuery(
            actor_id=actor, now=__import__("datetime").datetime.now(__import__("datetime").UTC)
        )
    )

    assert len(audit.events) == 3
    assert audit.events[0].id != audit.events[1].id  # type: ignore[attr-defined]
    assert audit.events[1].id != audit.events[2].id  # type: ignore[attr-defined]


async def test_missing_actor_raises_authorization_in_each_method() -> None:
    audit = _AuditLog()
    svc = _stack(audit, [_make_exp()])
    with pytest.raises(Exception, match="actor_id"):
        await svc.bandeja(BandejaQuery(actor_id=None, estado=None, codigo=None))
    with pytest.raises(Exception, match="actor_id"):
        await svc.tareas_calculadas(
            TareasCalculadasQuery(
                actor_id=None, now=__import__("datetime").datetime.now(__import__("datetime").UTC)
            )
        )
    with pytest.raises(Exception, match="actor_id"):
        await svc.exportar_excel(ExportacionExcelCommand(actor_id=None, filtros={}))
    assert audit.events == []
