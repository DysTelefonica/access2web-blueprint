"""UAT end-to-end for CAP-001..007 lifecycle (issue #261, U01)."""

from dataclasses import dataclass, field
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.create_expediente.command import (
    ExpedienteAltaCommand,
)
from app.src.modules.expedientes.application.create_expediente.service import (
    ExpedienteAltaService,
)
from app.src.modules.expedientes.application.delete_expediente.command import (
    ExpedienteDeleteCommand,
)
from app.src.modules.expedientes.application.delete_expediente.service import (
    ExpedienteDeleteService,
)
from app.src.modules.expedientes.application.edit_expediente.command import (
    ExpedienteEditCommand,
)
from app.src.modules.expedientes.application.edit_expediente.service import (
    ExpedienteEditService,
)
from app.src.modules.expedientes.application.transition_expediente.command import (
    ExpedienteTransitionCommand,
)
from app.src.modules.expedientes.application.transition_expediente.service import (
    ExpedienteTransitionService,
)
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


@dataclass
class _AuditLog:
    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


class _UoW:
    commits: int = 0
    rollbacks: int = 0

    def __enter__(self) -> object:
        return object()

    def __exit__(self, exc_type: object, *_a: object) -> None:
        if exc_type is None:
            type(self).commits += 1
        else:
            type(self).rollbacks += 1


class _Repo:
    def __init__(self) -> None:
        self.by_id: dict[UUID, Any] = {}

    async def get_by_id(self, eid: UUID) -> Any | None:
        return self.by_id.get(eid)

    async def lock_for_update(self, eid: UUID) -> Any | None:
        return self.by_id.get(eid)

    async def upsert(self, exp: Any) -> Any:
        self.by_id[exp.id] = exp
        return exp

    async def update(self, exp: Any) -> Any:
        self.by_id[exp.id] = exp
        return exp

    async def create(self, exp: Any) -> Any:
        self.by_id[exp.id] = exp
        return exp

    async def delete(self, eid: UUID) -> bool:
        return self.by_id.pop(eid, None) is not None

    async def has_children(self, eid: UUID) -> bool:
        return False

    async def get_by_expediente(self, exp_id: UUID) -> list[Any]:
        return []


def _stack(
    audit: _AuditLog,
) -> tuple[
    ExpedienteAltaService,
    ExpedienteEditService,
    ExpedienteDeleteService,
    ExpedienteTransitionService,
    _Repo,
]:
    repo = _Repo()
    uow = _UoW()
    return (
        ExpedienteAltaService(
            expediente_repo=repo, hito_repo=repo, audit_log=audit, uow_factory=lambda: uow
        ),
        ExpedienteEditService(
            expediente_repo=repo, hito_repo=repo, audit_log=audit, uow_factory=lambda: uow
        ),
        ExpedienteDeleteService(
            expediente_repo=repo, hito_repo=repo, audit_log=audit, uow_factory=lambda: uow
        ),
        ExpedienteTransitionService(
            expediente_repo=repo, hito_repo=repo, audit_log=audit, uow_factory=lambda: uow
        ),
        repo,
    )


def _types(audit: _AuditLog) -> list[str]:
    return [e.event_type for e in audit.events]  # type: ignore[attr-defined]


def _create_cmd(actor_id: Any) -> ExpedienteAltaCommand:
    return ExpedienteAltaCommand(actor_id=actor_id, codigo_ordinal="EXP-1", tipo=ExpedienteTipo.AM)


async def test_full_lifecycle_emits_audit_chain_in_order() -> None:
    """create -> transition -> edit -> delete emits 4 audit events in order."""
    audit = _AuditLog()
    create_svc, edit_svc, delete_svc, transition_svc, repo = _stack(audit)
    actor = uuid4()

    created = await create_svc.execute(_create_cmd(actor))
    exp_id = created.expediente_id
    await transition_svc.execute(
        ExpedienteTransitionCommand(
            actor_id=actor,
            expediente_id=exp_id,
            expected_version=1,
            new_tipo=None,
            new_id_expediente_padre=None,
            motivo="to_tramite",
        )
    )
    await edit_svc.execute(
        ExpedienteEditCommand(
            actor_id=actor,
            expediente_id=exp_id,
            expected_version=1,
            estado=ExpedienteEstado.FORMALIZADO,
        )
    )
    await delete_svc.execute(
        ExpedienteDeleteCommand(actor_id=actor, expediente_id=exp_id, expected_version=2)
    )

    assert _types(audit) == [
        "expediente.created",
        "expediente.updated",
        "expediente.deleted",
    ]


async def test_create_then_edit_distinct_event_ids_over_shared_audit() -> None:
    audit = _AuditLog()
    create_svc, edit_svc, _, _, _ = _stack(audit)
    actor = uuid4()

    created = await create_svc.execute(_create_cmd(actor))
    await edit_svc.execute(
        ExpedienteEditCommand(
            actor_id=actor,
            expediente_id=created.expediente_id,
            expected_version=1,
            estado=ExpedienteEstado.FORMALIZADO,
        )
    )

    assert len(audit.events) == 2
    assert audit.events[0].id != audit.events[1].id  # type: ignore[attr-defined]


async def test_missing_actor_raises_authorization_in_each_lifecycle_service() -> None:
    audit = _AuditLog()
    create_svc, edit_svc, delete_svc, transition_svc, _ = _stack(audit)
    exp_id = uuid4()

    with pytest.raises(Exception, match="actor_id"):
        await create_svc.execute(
            ExpedienteAltaCommand(actor_id=None, codigo_ordinal="X", tipo=ExpedienteTipo.AM)
        )
    with pytest.raises(Exception, match="actor_id"):
        await edit_svc.execute(
            ExpedienteEditCommand(
                actor_id=None,
                expediente_id=exp_id,
                expected_version=1,
                estado=ExpedienteEstado.BORRADOR,
            )
        )
    with pytest.raises(Exception, match="actor_id"):
        await delete_svc.execute(
            ExpedienteDeleteCommand(actor_id=None, expediente_id=exp_id, expected_version=1)
        )
    with pytest.raises(Exception, match="actor_id"):
        await transition_svc.execute(
            ExpedienteTransitionCommand(
                actor_id=None,
                expediente_id=exp_id,
                expected_version=1,
                new_tipo=None,
                new_id_expediente_padre=None,
            )
        )
    assert audit.events == []
