"""Strict TDD for EXP-CAP-045 Estado de sesión adapter (A03)."""

from __future__ import annotations

import dataclasses
import importlib
from dataclasses import dataclass, field
from datetime import datetime
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.session_state.command import (
    EXPEDIENTES_APP_ID,
    SessionState,
    SessionStateAuthorizationError,
    SessionStateCommand,
)
from app.src.modules.expedientes.application.session_state.service import (
    SessionStateService,
)


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)


def _make_service(audit: _Audit | None = None) -> tuple[SessionStateService, _Audit]:
    audit = audit or _Audit()
    return SessionStateService(audit_log=audit), audit


def _cmd(*, actor_id: UUID | None) -> SessionStateCommand:
    return SessionStateCommand(actor_id=actor_id)


async def test_bind_returns_session_state_for_valid_actor() -> None:
    actor = uuid4()
    service, audit = _make_service()
    state = await service.bind(_cmd(actor_id=actor))
    assert isinstance(state, SessionState)
    assert state.actor_id == actor
    assert isinstance(state.session_id, UUID)
    assert state.app_id == EXPEDIENTES_APP_ID
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "authz.session.bound"


async def test_get_current_returns_bound_state() -> None:
    actor = uuid4()
    service, _ = _make_service()
    bound = await service.bind(_cmd(actor_id=actor))
    assert service.get_current() is bound


async def test_get_current_returns_none_when_unbound() -> None:
    service, audit = _make_service()
    assert service.get_current() is None
    assert audit.events == []


async def test_concurrent_binds_produce_independent_states() -> None:
    service_a, _ = _make_service()
    service_b, _ = _make_service()
    state_a = await service_a.bind(_cmd(actor_id=uuid4()))
    state_b = await service_b.bind(_cmd(actor_id=uuid4()))
    assert service_a.get_current() is state_a
    assert service_b.get_current() is state_b
    assert service_a.get_current() is not state_b
    assert service_b.get_current() is not state_a


async def test_deny_when_actor_id_missing() -> None:
    service, audit = _make_service()
    with pytest.raises(SessionStateAuthorizationError, match="actor_id"):
        await service.bind(_cmd(actor_id=None))
    assert audit.events == []


async def test_deny_when_actor_id_empty_uuid() -> None:
    service, audit = _make_service()
    with pytest.raises(SessionStateAuthorizationError, match="actor_id"):
        await service.bind(_cmd(actor_id=UUID(int=0)))
    assert audit.events == []


async def test_audit_event_carries_actor_session_and_capability() -> None:
    actor = uuid4()
    service, audit = _make_service()
    state = await service.bind(_cmd(actor_id=actor))
    event = audit.events[0]
    assert event.actor_id == actor
    assert event.target_id == actor
    assert event.capacidad == "EXP-CAP-045"
    assert event.payload["session_id"] == state.session_id


async def test_clear_resets_state_and_emits_event() -> None:
    actor = uuid4()
    service, audit = _make_service()
    state = await service.bind(_cmd(actor_id=actor))
    await service.clear()
    assert service.get_current() is None
    assert len(audit.events) == 2
    assert audit.events[0].event_type == "authz.session.bound"
    assert audit.events[1].event_type == "authz.session.cleared"
    assert audit.events[1].payload["session_id"] == state.session_id


async def test_clear_when_unbound_is_idempotent_no_audit() -> None:
    service, audit = _make_service()
    await service.clear()
    assert service.get_current() is None
    assert audit.events == []


async def test_session_state_is_immutable() -> None:
    state = SessionState(
        actor_id=uuid4(),
        session_id=uuid4(),
        bound_at=datetime(2026, 1, 1),
        app_id=EXPEDIENTES_APP_ID,
    )
    with pytest.raises(dataclasses.FrozenInstanceError):
        state.actor_id = uuid4()  # type: ignore[misc]


async def test_no_mutable_globals_in_session_state_module() -> None:
    service, _ = _make_service()
    await service.bind(_cmd(actor_id=uuid4()))
    module = importlib.import_module(
        "app.src.modules.expedientes.application.session_state.service"
    )
    forbidden_keys = {"events", "state", "_state_value", "current"}
    for name, value in vars(module).items():
        if (
            name.startswith("__")
            or isinstance(value, (type, SessionStateService))
            or callable(value)
        ):
            continue
        assert name not in forbidden_keys, (
            f"module-level {name!r} in session_state.service is per-request mutable"
        )
