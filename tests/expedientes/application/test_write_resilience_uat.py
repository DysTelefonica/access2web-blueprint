"""UAT end-to-end for CAP-030..032 write-resilience (issue #278, U05)."""

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.autosave_general.command import AutosaveGeneralCommand
from app.src.modules.expedientes.application.autosave_general.service import AutosaveGeneralService
from app.src.modules.expedientes.application.autosave_related.command import AutosaveRelatedCommand
from app.src.modules.expedientes.application.autosave_related.service import AutosaveRelatedService
from app.src.modules.expedientes.application.feedback_guard.command import FeedbackGuardCommand
from app.src.modules.expedientes.application.feedback_guard.service import FeedbackGuardService
from app.src.modules.expedientes.domain.anualidad import Anualidad
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo
from app.src.modules.expedientes.domain.hito import Hito, HitoEstado
from app.src.modules.expedientes.domain.modificado import Modificado
from app.src.modules.expedientes.ports.idempotency import IdempotencyRecord


@dataclass
class _AuditLog:
    events: list[object] = field(default_factory=list)

    async def append(self, e: object) -> None:
        self.events.append(e)

    async def record_change(self, c: object) -> None:
        return None


class _UoW:
    def __enter__(self) -> object:
        return object()

    def __exit__(self, *_a: object) -> None:
        return None


class _Repo:
    def __init__(self, expediente: object) -> None:
        self._e = expediente

    async def get_by_id(self, eid: object, *_a: object, **_kw: object) -> object:
        return self._e if eid == self._e.id else None

    async def upsert(self, *_a: object, **_kw: object) -> object:
        return self._e

    async def update(self, *_a: object, **_kw: object) -> object:
        return self._e

    async def get_by_expediente(self, *_a: object, **_kw: object) -> object:
        return []


class _Idem:
    def __init__(self) -> None:
        self.store: dict[UUID, IdempotencyRecord] = {}

    async def get(self, key: UUID) -> IdempotencyRecord | None:
        return self.store.get(key)

    async def record(self, record: IdempotencyRecord) -> None:
        self.store[record.key] = record


def _stack(
    audit: _AuditLog, idem: _Idem, expediente: object
) -> tuple[AutosaveGeneralService, AutosaveRelatedService, FeedbackGuardService]:
    repo, uow = _Repo(expediente), lambda: _UoW()
    return (
        AutosaveGeneralService(
            expediente_repo=repo, idempotency=idem, audit_log=audit, uow_factory=uow
        ),
        AutosaveRelatedService(
            expediente_repo=repo,
            hito_repo=repo,
            modificado_repo=repo,
            anualidad_repo=repo,
            idempotency=idem,
            audit_log=audit,
            uow_factory=uow,
        ),
        FeedbackGuardService(idempotency=idem, uow_factory=uow, audit_log=audit),
    )


def _make_exp(exp_id: UUID) -> object:
    return Expediente(
        id=exp_id, tipo=ExpedienteTipo.AM, estado=ExpedienteEstado.BORRADOR, version=1
    )


_T = [
    "expediente.general.autosaved",
    "hito.autosaved",
    "modificado.autosaved",
    "anualidad.autosaved",
    "feedback.busy",
]
_C = ["EXP-CAP-030", "EXP-CAP-031", "EXP-CAP-031", "EXP-CAP-031", "EXP-CAP-032"]


async def test_full_chain_emits_five_audit_events_in_order() -> None:
    audit, idem, actor, exp_id = _AuditLog(), _Idem(), uuid4(), uuid4()
    g, r, f = _stack(audit, idem, _make_exp(exp_id))

    await g.execute(
        AutosaveGeneralCommand(
            expediente_id=exp_id,
            expected_version=1,
            idempotency_key=uuid4(),
            fecha_inicio_contrato=date(2026, 1, 1),
            fecha_fin_contrato=None,
            actor_id=actor,
        )
    )
    await r.execute(
        AutosaveRelatedCommand(
            idempotency_key=uuid4(),
            expediente_id=exp_id,
            expected_version=1,
            hitos=(
                Hito(
                    id=uuid4(),
                    id_expediente=exp_id,
                    fecha_hito=date(2026, 2, 1),
                    garantia_fecha_fin=None,
                    estado=HitoEstado.PENDIENTE,
                ),
            ),
            modificados=(
                Modificado(
                    id=uuid4(),
                    id_expediente=exp_id,
                    n_modificado="M-1",
                    fecha_firma=date(2026, 3, 1),
                    fecha_fin=date(2026, 6, 1),
                    descripcion="x",
                ),
            ),
            anualidades=(
                Anualidad(
                    id=uuid4(),
                    id_expediente=exp_id,
                    anio=2026,
                    bi_iva=Decimal("1"),
                    bi_ipsi=None,
                    bi_igic=None,
                    bi_exenta=Decimal("0"),
                    iva=Decimal("0"),
                    ipsi=None,
                    igic=None,
                    periodo_facturacion="Q1",
                ),
            ),
            actor_id=actor,
        )
    )
    await f.execute(FeedbackGuardCommand(idempotency_key=uuid4(), operation="edit", actor_id=actor))

    assert [e.event_type for e in audit.events] == _T  # type: ignore[attr-defined]
    assert [e.capacidad for e in audit.events] == _C  # type: ignore[attr-defined]


async def test_idempotent_general_autosave_does_not_double_emit() -> None:
    audit, idem = _AuditLog(), _Idem()
    actor, exp_id, key = uuid4(), uuid4(), uuid4()
    g, _, _ = _stack(audit, idem, _make_exp(exp_id))

    cmd = AutosaveGeneralCommand(
        expediente_id=exp_id,
        expected_version=1,
        idempotency_key=key,
        fecha_inicio_contrato=date(2026, 1, 1),
        fecha_fin_contrato=None,
        actor_id=actor,
    )
    await g.execute(cmd)
    await g.execute(cmd)

    assert len(audit.events) == 1


async def test_missing_actor_raises_authorization_in_each_service() -> None:
    audit, idem, exp_id = _AuditLog(), _Idem(), uuid4()
    g, r, f = _stack(audit, idem, _make_exp(exp_id))
    cmds = [
        (
            "g",
            g.execute,
            AutosaveGeneralCommand(
                expediente_id=exp_id,
                expected_version=1,
                idempotency_key=uuid4(),
                fecha_inicio_contrato=None,
                fecha_fin_contrato=None,
                actor_id=None,
            ),
        ),
        (
            "r",
            r.execute,
            AutosaveRelatedCommand(
                idempotency_key=uuid4(),
                expediente_id=exp_id,
                expected_version=1,
                hitos=(),
                modificados=(),
                anualidades=(),
                actor_id=None,
            ),
        ),
        (
            "f",
            f.execute,
            FeedbackGuardCommand(idempotency_key=uuid4(), operation="edit", actor_id=None),
        ),
    ]
    for _name, method, cmd in cmds:
        with pytest.raises(Exception, match="actor_id"):
            await method(cmd)
    assert audit.events == []
