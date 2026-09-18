"""UAT end-to-end for CAP-008..014 related-data vertical (issue #262, U02)."""

from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal
from uuid import uuid4

import pytest

from app.src.modules.expedientes.application.anexos.command import CreateAnexoCommand
from app.src.modules.expedientes.application.anexos.service import AnexosService
from app.src.modules.expedientes.application.register_anualidad.command import (
    RegisterAnualidadCommand,
)
from app.src.modules.expedientes.application.register_anualidad.service import (
    RegisterAnualidadService,
)
from app.src.modules.expedientes.application.register_hito.command import RegisterHitoCommand
from app.src.modules.expedientes.application.register_hito.service import RegisterHitoService
from app.src.modules.expedientes.application.register_juridica.command import (
    RegisterJuridicaCommand,
)
from app.src.modules.expedientes.application.register_juridica.service import (
    RegisterJuridicaService,
)
from app.src.modules.expedientes.application.register_modificado.command import (
    RegisterModificadoCommand,
)
from app.src.modules.expedientes.application.register_modificado.service import (
    RegisterModificadoService,
)
from app.src.modules.expedientes.application.register_responsable.command import (
    RegisterResponsableCommand,
)
from app.src.modules.expedientes.application.register_responsable.service import (
    RegisterResponsableService,
)
from app.src.modules.expedientes.application.register_suministrador.command import (
    RegisterSuministradorCommand,
)
from app.src.modules.expedientes.application.register_suministrador.service import (
    RegisterSuministradorService,
)
from app.src.modules.expedientes.domain.anexo.reference import AnexoReference
from app.src.modules.expedientes.domain.anexo.retention import RetentionLevel, RetentionPolicy
from app.src.modules.expedientes.domain.anexo.size_limit import SizeLimitPolicy
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


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

    async def upsert(self, *_a: object, **_kw: object) -> None:
        return None

    async def add(self, *_a: object, **_kw: object) -> None:
        return None

    async def delete(self, *_a: object, **_kw: object) -> None:
        return None

    async def get_by_expediente(self, *_a: object, **_kw: object) -> object:
        return None

    async def list_active(self, *_a: object, **_kw: object) -> list[object]:
        return []

    async def lock_for_update(self, *_a: object, **_kw: object) -> None:
        return None


_SIZE = SizeLimitPolicy(max_size_bytes=10**8)
_RET = RetentionPolicy(level=RetentionLevel.OPERATIONAL, duration=timedelta(days=30))


def _stack(audit: _AuditLog, exp: object) -> dict[str, object]:
    repo, uow = _Repo(exp), lambda: _UoW()
    return {
        "hito": RegisterHitoService(
            expediente_repo=repo, hito_repo=repo, audit_log=audit, uow_factory=uow
        ),
        "modificado": RegisterModificadoService(
            expediente_repo=repo, modificado_repo=repo, audit_log=audit, uow_factory=uow
        ),
        "anexo": AnexosService(
            expediente_repo=repo,
            anexo_repo=repo,
            audit_log=audit,
            size_limit=_SIZE,
            uow_factory=uow,
        ),
        "anualidad": RegisterAnualidadService(
            expediente_repo=repo, anualidad_repo=repo, audit_log=audit, uow_factory=uow
        ),
        "responsable": RegisterResponsableService(
            expediente_repo=repo, responsable_repo=repo, audit_log=audit, uow_factory=uow
        ),
        "juridica": RegisterJuridicaService(
            expediente_repo=repo, juridica_repo=repo, audit_log=audit, uow_factory=uow
        ),
        "suministrador": RegisterSuministradorService(
            expediente_repo=repo, suministrador_repo=repo, audit_log=audit, uow_factory=uow
        ),
    }


def _make_exp(exp_id: object) -> object:
    return Expediente(
        id=exp_id, tipo=ExpedienteTipo.AM, estado=ExpedienteEstado.BORRADOR, version=1
    )


_T = [
    "hito.registered",
    "modificado.registered",
    "anexo.created",
    "anualidad.registered",
    "responsable.registered",
    "juridica.registered",
    "suministrador.registered",
]
_C = [f"EXP-CAP-{n:03d}" for n in (8, 9, 10, 11, 12, 13, 14)]


async def test_full_chain_emits_seven_audit_events_in_order() -> None:
    audit, actor, exp_id = _AuditLog(), uuid4(), uuid4()
    s = _stack(audit, _make_exp(exp_id))

    await s["hito"].execute(
        RegisterHitoCommand(
            hito_id=uuid4(),
            expediente_id=exp_id,
            descripcion="hito",
            fecha_hito=date(2026, 1, 15),
            garantia_fecha_fin=date(2027, 1, 15),
            importe=Decimal("100"),
            actor_id=actor,
        )
    )
    await s["modificado"].execute(
        RegisterModificadoCommand(
            modificado_id=uuid4(),
            expediente_id=exp_id,
            n_modificado="MOD-1",
            fecha_firma=date(2026, 6, 1),
            fecha_fin=date(2026, 12, 1),
            descripcion="x",
            actor_id=actor,
        )
    )
    await s["anexo"].create(
        CreateAnexoCommand(
            anexo_id=uuid4(),
            expediente_id=exp_id,
            nombre="doc.pdf",
            referencia=AnexoReference(
                storage_ref=f"storage://exp/{uuid4()}",
                content_type="application/pdf",
                size_bytes=1024,
            ),
            retention=_RET,
            actor_id=actor,
        )
    )
    await s["anualidad"].execute(
        RegisterAnualidadCommand(
            anualidad_id=uuid4(),
            expediente_id=exp_id,
            anio=2026,
            bi_iva=Decimal("1000"),
            bi_ipsi=None,
            bi_igic=None,
            bi_exenta=Decimal("0"),
            iva=Decimal("210"),
            ipsi=None,
            igic=None,
            periodo_facturacion="Q1",
            actor_id=actor,
        )
    )
    await s["responsable"].execute(
        RegisterResponsableCommand(
            responsable_id=uuid4(),
            expediente_id=exp_id,
            usuario_id=uuid4(),
            rol="JP",
            correo_siempre=True,
            es_jefe_proyecto=True,
            es_preventa=False,
            actor_id=actor,
        )
    )
    await s["juridica"].execute(
        RegisterJuridicaCommand(
            juridica_id=uuid4(),
            expediente_id=exp_id,
            id_juridica=uuid4(),
            id_suministrador=None,
            contratista_principal=True,
            sub_contratista=False,
            actor_id=actor,
        )
    )
    await s["suministrador"].execute(
        RegisterSuministradorCommand(
            suministrador_id=uuid4(),
            expediente_id=exp_id,
            id_suministrador=uuid4(),
            id_padre=None,
            descripcion="UTE",
            contratista_principal=True,
            sub_contratista=False,
            actor_id=actor,
        )
    )

    assert [e.event_type for e in audit.events] == _T  # type: ignore[attr-defined]
    assert [e.capacidad for e in audit.events] == _C  # type: ignore[attr-defined]


async def test_two_services_over_same_audit_emit_distinct_event_ids() -> None:
    audit, actor_a, actor_b = _AuditLog(), uuid4(), uuid4()
    exp_a, exp_b = uuid4(), uuid4()
    s_a = _stack(audit, _make_exp(exp_a))
    await s_a["hito"].execute(
        RegisterHitoCommand(
            hito_id=uuid4(),
            expediente_id=exp_a,
            descripcion="A",
            fecha_hito=date(2026, 1, 1),
            garantia_fecha_fin=None,
            importe=None,
            actor_id=actor_a,
        )
    )
    s_b = _stack(audit, _make_exp(exp_b))
    await s_b["modificado"].execute(
        RegisterModificadoCommand(
            modificado_id=uuid4(),
            expediente_id=exp_b,
            n_modificado="M-1",
            fecha_firma=date(2026, 6, 1),
            fecha_fin=date(2026, 12, 1),
            descripcion="x",
            actor_id=actor_b,
        )
    )

    assert [e.actor_id for e in audit.events] == [actor_a, actor_b]  # type: ignore[attr-defined]
    assert audit.events[0].id != audit.events[1].id  # type: ignore[attr-defined]


async def test_missing_actor_raises_authorization_in_each_service() -> None:
    audit, exp_id = _AuditLog(), uuid4()
    s = _stack(audit, _make_exp(exp_id))
    cmds = [
        (
            "hito",
            "execute",
            RegisterHitoCommand(
                hito_id=uuid4(),
                expediente_id=exp_id,
                descripcion="x",
                fecha_hito=date(2026, 1, 1),
                garantia_fecha_fin=None,
                importe=None,
                actor_id=None,
            ),
        ),
        (
            "modificado",
            "execute",
            RegisterModificadoCommand(
                modificado_id=uuid4(),
                expediente_id=exp_id,
                n_modificado="M",
                fecha_firma=date(2026, 6, 1),
                fecha_fin=date(2026, 12, 1),
                descripcion="x",
                actor_id=None,
            ),
        ),
        (
            "anexo",
            "create",
            CreateAnexoCommand(
                anexo_id=uuid4(),
                expediente_id=exp_id,
                nombre="x",
                referencia=AnexoReference(
                    storage_ref=f"storage://exp/{uuid4()}",
                    content_type="application/pdf",
                    size_bytes=1024,
                ),
                retention=_RET,
                actor_id=None,
            ),
        ),
    ]
    for name, method, cmd in cmds:
        with pytest.raises(Exception, match="actor_id"):
            await getattr(s[name], method)(cmd)
    assert audit.events == []
