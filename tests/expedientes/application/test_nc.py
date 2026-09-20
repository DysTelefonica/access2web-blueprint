"""Strict TDD for EXP-CAP-054 NC adapter (H04, issue #251)."""

from dataclasses import dataclass, field
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.nc.command import (
    NCAuthorizationError,
    NCDependencyError,
    NCLookupCommand,
)
from app.src.modules.expedientes.application.nc.service import NCService
from app.src.modules.expedientes.ports.nc import NCLookup, NCRecord


@dataclass
class _AuditLog:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)


class _FakeNC:
    def __init__(self) -> None:
        self.by_exp: dict[UUID, Any] = {}
        self.exc: BaseException | None = None

    async def lookup(self, request: NCLookup, credential: str) -> Any:
        if self.exc is not None:
            raise self.exc
        if request.expediente_id is None:
            return None
        return self.by_exp.get(request.expediente_id)


def _service() -> tuple[NCService, _AuditLog, _FakeNC]:
    audit = _AuditLog()
    fake = _FakeNC()
    return NCService(nc=fake, audit_log=audit), audit, fake


async def test_lookup_by_expediente_emits_one_event_with_estado() -> None:
    svc, audit, fake = _service()
    actor, exp_id, nc_id = uuid4(), uuid4(), uuid4()
    fake.by_exp[exp_id] = NCRecord(nc_id=nc_id, expediente_id=exp_id, estado="abierta", payload={})

    result = await svc.lookup(NCLookupCommand(actor_id=actor, expediente_id=exp_id))

    assert result.estado == "abierta"
    assert result.expediente_id == exp_id
    assert audit.events[0].event_type == "nc.lookup.recorded"  # type: ignore[attr-defined]
    assert audit.events[0].capacidad == "EXP-CAP-054"  # type: ignore[attr-defined]


async def test_lookup_without_expediente_and_s4h_raises_authorization() -> None:
    svc, audit, _ = _service()
    with pytest.raises(NCAuthorizationError, match="expediente_id or s4h_code"):
        await svc.lookup(NCLookupCommand(actor_id=uuid4()))
    assert audit.events == []


async def test_missing_actor_raises_authorization() -> None:
    svc, audit, _ = _service()
    with pytest.raises(NCAuthorizationError, match="actor_id"):
        await svc.lookup(NCLookupCommand(actor_id=None, expediente_id=uuid4()))
    assert audit.events == []


async def test_dependency_failure_wraps_underlying_exception() -> None:
    fake = _FakeNC()
    fake.exc = RuntimeError("nc down")
    svc = NCService(nc=fake, audit_log=_AuditLog())
    with pytest.raises(NCDependencyError) as excinfo:
        await svc.lookup(NCLookupCommand(actor_id=uuid4(), expediente_id=uuid4()))
    assert isinstance(excinfo.value.__cause__, RuntimeError)
