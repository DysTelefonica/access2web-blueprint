"""Strict TDD for EXP-CAP-044 CurrentPrincipal adapter (A02)."""

from __future__ import annotations

from dataclasses import dataclass, field
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.current_principal.command import (
    CurrentPrincipalAuthorizationError,
    CurrentPrincipalCommand,
    EXPEDIENTES_APP_ID,
    Principal,
)
from app.src.modules.expedientes.application.current_principal.service import (
    CurrentPrincipalService,
)


@dataclass
class _FakeLanzadera:
    principal_to_return: object = None
    exception_to_raise: BaseException | None = None

    async def get_principal(self, actor_id: UUID) -> object:
        if self.exception_to_raise is not None:
            raise self.exception_to_raise
        return self.principal_to_return


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)


def _make_service(
    *,
    lanzadera: _FakeLanzadera | None = None,
    audit: _Audit | None = None,
) -> tuple[CurrentPrincipalService, _Audit]:
    audit = audit or _Audit()
    lanzadera = lanzadera or _FakeLanzadera()
    service = CurrentPrincipalService(lanzadera_principal=lanzadera, audit_log=audit)
    return service, audit


def _command(*, actor_id: UUID | None) -> CurrentPrincipalCommand:
    return CurrentPrincipalCommand(actor_id=actor_id)


async def test_loaded_returns_principal_with_permissions() -> None:
    actor = uuid4()
    principal = Principal(
        user_id=actor, app_id=EXPEDIENTES_APP_ID, permissions=frozenset({"e2e.session"})
    )
    lanzadera = _FakeLanzadera(principal_to_return=principal)
    service, audit = _make_service(lanzadera=lanzadera)

    result = await service.get_current(_command(actor_id=actor))

    assert result.loaded is True
    assert result.principal.permissions == frozenset({"e2e.session"})
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "authz.principal.loaded"


async def test_deny_when_lanzadera_returns_principal_with_no_permissions() -> None:
    actor = uuid4()
    principal = Principal(user_id=actor, app_id=EXPEDIENTES_APP_ID, permissions=frozenset())
    lanzadera = _FakeLanzadera(principal_to_return=principal)
    service, audit = _make_service(lanzadera=lanzadera)

    result = await service.get_current(_command(actor_id=actor))

    assert result.loaded is False
    assert result.principal.permissions == frozenset()
    assert audit.events[0].event_type == "authz.principal.failed"
    assert audit.events[0].payload["reason"] == "empty_permissions"


async def test_deny_when_lanzadera_returns_none() -> None:
    actor = uuid4()
    lanzadera = _FakeLanzadera(principal_to_return=None)
    service, audit = _make_service(lanzadera=lanzadera)

    result = await service.get_current(_command(actor_id=actor))

    assert result.loaded is False
    assert audit.events[0].payload["reason"] == "missing_principal"


async def test_deny_when_lanzadera_raises_runtime_error() -> None:
    actor = uuid4()
    lanzadera = _FakeLanzadera(exception_to_raise=RuntimeError("boom"))
    service, audit = _make_service(lanzadera=lanzadera)

    result = await service.get_current(_command(actor_id=actor))

    assert result.loaded is False
    assert audit.events[0].payload["reason"] == "lanzadera_error"
    assert audit.events[0].payload["error_class"] == "RuntimeError"


async def test_raises_on_missing_actor_id() -> None:
    service, _ = _make_service()

    with pytest.raises(CurrentPrincipalAuthorizationError, match="actor_id"):
        await service.get_current(_command(actor_id=None))


async def test_raises_on_empty_actor_id() -> None:
    service, _ = _make_service()

    with pytest.raises(CurrentPrincipalAuthorizationError):
        await service.get_current(_command(actor_id=UUID(int=0)))


async def test_loaded_audit_event_carries_actor_and_target() -> None:
    actor = uuid4()
    principal = Principal(
        user_id=actor, app_id=EXPEDIENTES_APP_ID, permissions=frozenset({"e2e.session"})
    )
    lanzadera = _FakeLanzadera(principal_to_return=principal)
    service, audit = _make_service(lanzadera=lanzadera)

    await service.get_current(_command(actor_id=actor))

    assert audit.events[0].actor_id == actor
    assert audit.events[0].target_id == principal.user_id
    assert audit.events[0].payload["app_id"] == EXPEDIENTES_APP_ID


async def test_failed_audit_event_records_error_class() -> None:
    actor = uuid4()
    lanzadera = _FakeLanzadera(exception_to_raise=ConnectionError("timeout"))
    service, audit = _make_service(lanzadera=lanzadera)

    await service.get_current(_command(actor_id=actor))

    assert audit.events[0].payload["error_class"] == "ConnectionError"
