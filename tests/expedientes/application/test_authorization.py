"""Strict TDD for EXP-CAP-043 authorization policy (A01)."""

from __future__ import annotations

from dataclasses import dataclass, field
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.authorization.command import (
    EFFECTIVE_DENY,
    AuthorizationAuthorizationError,
    AuthorizationCommand,
    AuthorizationValidationError,
)
from app.src.modules.expedientes.application.authorization.service import (
    AuthorizationService,
)


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


def _make_service() -> tuple[AuthorizationService, _Audit]:
    audit = _Audit()
    service = AuthorizationService(audit_log=audit)
    return service, audit


def _command(
    capability: str = "e2e.session",
    *,
    actor_id: UUID | None = None,
    permissions: set[str] | None = None,
) -> AuthorizationCommand:
    return AuthorizationCommand(
        actor_id=actor_id,
        capability=capability,
        permissions=permissions or set(),
    )


async def test_allow_when_actor_has_capability() -> None:
    service, audit = _make_service()

    result = await service.authorize(
        _command(capability="e2e.session", permissions={"e2e.session"}, actor_id=uuid4())
    )

    assert result.allowed is True
    assert result.denied_reason is None
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "e2e.authz.allowed"
    assert audit.events[0].capacidad == "e2e.session"


async def test_deny_when_actor_lacks_capability() -> None:
    service, audit = _make_service()

    result = await service.authorize(
        _command(capability="e2e.session", permissions=set(), actor_id=uuid4())
    )

    assert result.allowed is False
    assert result.denied_reason == EFFECTIVE_DENY
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "e2e.authz.denied"


async def test_deny_returns_deny_by_default_constant_when_no_permissions() -> None:
    service, _ = _make_service()

    result = await service.authorize(
        _command(capability="bandeja.read", permissions=set(), actor_id=uuid4())
    )

    assert result.allowed is False
    assert result.denied_reason == EFFECTIVE_DENY


async def test_deny_without_actor() -> None:
    service, _ = _make_service()

    with pytest.raises(AuthorizationAuthorizationError, match="actor_id"):
        await service.authorize(_command(capability="e2e.session", actor_id=None))


async def test_deny_with_invalid_capability() -> None:
    service, _ = _make_service()

    with pytest.raises(AuthorizationValidationError, match="capability"):
        await service.authorize(
            _command(capability="", permissions={"e2e.session"}, actor_id=uuid4())
        )


async def test_deny_with_empty_capability_string() -> None:
    service, _ = _make_service()

    with pytest.raises(AuthorizationValidationError, match="capability"):
        await service.authorize(
            _command(capability="   ", permissions={"e2e.session"}, actor_id=uuid4())
        )


async def test_deny_with_empty_actor_id() -> None:
    service, _ = _make_service()

    with pytest.raises(AuthorizationAuthorizationError):
        await service.authorize(_command(capability="e2e.session", actor_id=UUID(int=0)))


async def test_allow_audit_event_carries_actor_id() -> None:
    service, audit = _make_service()
    actor = uuid4()

    result = await service.authorize(
        _command(capability="e2e.session", actor_id=actor, permissions={"e2e.session"})
    )

    assert audit.events[0].actor_id == actor
    assert result.actor_id == actor


async def test_deny_audit_event_carries_actor_id() -> None:
    service, audit = _make_service()
    actor = uuid4()

    result = await service.authorize(
        _command(capability="e2e.session", actor_id=actor, permissions=set())
    )

    assert audit.events[0].actor_id == actor
    assert result.actor_id == actor
    assert result.allowed is False
