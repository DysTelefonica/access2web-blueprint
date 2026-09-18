"""UAT end-to-end for CAP-033..042 E2E-export (issue #279, U06)."""

from dataclasses import dataclass, field
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.batch_e2e.command import (
    BatchDetailInput,
    BatchE2ECommand,
)
from app.src.modules.expedientes.application.batch_e2e.service import BatchE2EService
from app.src.modules.expedientes.application.e2e_dto.command import E2EDtoCommand
from app.src.modules.expedientes.application.e2e_dto.service import E2EDtoService
from app.src.modules.expedientes.application.json_canonical.service import JsonCanonicalService
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


@dataclass
class _AuditLog:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)


class _NoopProcessor:
    async def process(self, _detail_id: UUID) -> None:
        return None


def _stack(audit: _AuditLog) -> dict[str, object]:
    perms = {"e2e.exchange", "e2e.export", "e2e.batch"}
    return {
        "dto": E2EDtoService(audit_log=audit, permissions=perms),
        "json": JsonCanonicalService(audit_log=audit, permissions=perms),
        "batch": BatchE2EService(audit_log=audit, processor=_NoopProcessor(), permissions=perms),
    }


def _make_exp() -> Expediente:
    return Expediente(
        id=uuid4(), tipo=ExpedienteTipo.AM, estado=ExpedienteEstado.BORRADOR, version=1
    )


def _types(audit: _AuditLog) -> list[str]:
    return [e.event_type for e in audit.events]  # type: ignore[attr-defined]


def _caps(audit: _AuditLog) -> list[str]:
    return [e.capacidad for e in audit.events]  # type: ignore[attr-defined]


async def test_dto_to_json_to_batch_emits_three_events_in_order() -> None:
    audit, actor, exp = _AuditLog(), uuid4(), _make_exp()
    s = _stack(audit)
    exped = [{"id": str(exp.id)}]

    await s["dto"].exchange(E2EDtoCommand(actor_id=actor, expediente=exp))
    await s["json"].render(collections={"expedientes": exped}, actor_id=actor)
    await s["batch"].execute(
        BatchE2ECommand(
            actor_id=actor,
            details=[BatchDetailInput(id=exp.id, payload={"action": "compute_hash"})],
            atomicity="all_or_nothing",
        )
    )

    assert _types(audit) == ["e2e.dto.exchanged", "e2e.canonical.rendered", "e2e.batch.executed"]
    assert _caps(audit) == ["EXP-CAP-033", "EXP-CAP-034", "EXP-CAP-035"]


async def test_two_services_over_same_audit_emit_distinct_event_ids() -> None:
    audit, actor, exp = _AuditLog(), uuid4(), _make_exp()
    s = _stack(audit)

    await s["dto"].exchange(E2EDtoCommand(actor_id=actor, expediente=exp))
    await s["json"].render(collections={"expedientes": [{"id": str(exp.id)}]}, actor_id=actor)

    assert len(audit.events) == 2
    assert audit.events[0].id != audit.events[1].id  # type: ignore[attr-defined]


async def test_missing_actor_raises_authorization_in_each_service() -> None:
    audit, exp = _AuditLog(), _make_exp()
    s = _stack(audit)

    with pytest.raises(Exception, match="actor_id"):
        await s["dto"].exchange(E2EDtoCommand(actor_id=None, expediente=exp))
    with pytest.raises(Exception, match="actor_id"):
        await s["json"].render(collections={}, actor_id=None)
    with pytest.raises(Exception, match="actor_id"):
        await s["batch"].execute(
            BatchE2ECommand(
                actor_id=None,
                details=[BatchDetailInput(id=exp.id, payload={"x": "y"})],
                atomicity="all_or_nothing",
            )
        )
    assert audit.events == []
