"""Strict TDD for EXP-CAP-053 Riesgos adapter (H03, issue #250)."""

from dataclasses import dataclass, field
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.riesgos.command import (
    RiesgosAuthorizationError,
    RiesgosDependencyError,
    RiesgosLookupCommand,
)
from app.src.modules.expedientes.application.riesgos.service import (
    RiesgosProject,
    RiesgosService,
)
from app.src.modules.expedientes.ports.riesgos import RiesgosLookup


@dataclass
class _AuditLog:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)


class _FakeRiesgos:
    """Returns a single Riesgos project keyed by expediente_id."""

    def __init__(self) -> None:
        self.by_id: dict[UUID, Any] = {}
        self.exc: BaseException | None = None

    async def lookup(self, request: RiesgosLookup, credential: str) -> Any:
        if self.exc is not None:
            raise self.exc
        return self.by_id.get(request.expediente_id)


def _service() -> tuple[RiesgosService, _AuditLog, _FakeRiesgos]:
    audit = _AuditLog()
    fake = _FakeRiesgos()
    return RiesgosService(riesgos=fake, audit_log=audit), audit, fake


async def test_lookup_emits_one_event_with_riesgos_id() -> None:
    svc, audit, fake = _service()
    actor, exp_id, riesgos_id = uuid4(), uuid4(), uuid4()
    from types import SimpleNamespace

    fake.by_id[exp_id] = SimpleNamespace(
        riesgos_id=riesgos_id, expediente_id=exp_id, status="open", payload={"risk": "low"}
    )

    result = await svc.lookup(RiesgosLookupCommand(actor_id=actor, expediente_id=exp_id))

    assert isinstance(result, RiesgosProject)
    assert result.riesgos_id == riesgos_id
    assert result.expediente_id == exp_id
    assert audit.events[0].event_type == "riesgos.lookup.recorded"  # type: ignore[attr-defined]
    assert audit.events[0].capacidad == "EXP-CAP-053"  # type: ignore[attr-defined]


async def test_missing_actor_raises_authorization_and_skips_audit() -> None:
    svc, audit, _ = _service()
    with pytest.raises(RiesgosAuthorizationError, match="actor_id"):
        await svc.lookup(RiesgosLookupCommand(actor_id=None, expediente_id=uuid4()))
    assert audit.events == []


async def test_dependency_failure_wraps_underlying_exception() -> None:
    fake = _FakeRiesgos()
    fake.exc = RuntimeError("down")
    svc = RiesgosService(riesgos=fake, audit_log=_AuditLog())
    with pytest.raises(RiesgosDependencyError) as excinfo:
        await svc.lookup(RiesgosLookupCommand(actor_id=uuid4(), expediente_id=uuid4()))
    assert isinstance(excinfo.value.__cause__, RuntimeError)


async def test_audit_payload_includes_riesgos_id_and_status() -> None:
    svc, audit, fake = _service()
    actor, exp_id, riesgos_id = uuid4(), uuid4(), uuid4()
    from types import SimpleNamespace

    fake.by_id[exp_id] = SimpleNamespace(
        riesgos_id=riesgos_id, expediente_id=exp_id, status="in_progress", payload={"risk": "high"}
    )

    await svc.lookup(RiesgosLookupCommand(actor_id=actor, expediente_id=exp_id))

    event = audit.events[0]  # type: ignore[attr-defined]
    assert event.payload["riesgos_id"] == str(riesgos_id)
    assert event.payload["status"] == "in_progress"
    assert event.payload["app_id"] == 19  # EXPEDIENTES_APP_ID
