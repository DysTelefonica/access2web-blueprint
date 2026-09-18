"""Strict TDD for EXP-CAP-046 Auditoría adapter (A04)."""

from __future__ import annotations

import dataclasses
import inspect
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.audit._evidence import EXPEDIENTES_APP_ID
from app.src.modules.expedientes.application.audit.command import (
    AuditAuthorizationError,
    AuditCommand,
    AuditDependencyError,
    AuditResult,
)
from app.src.modules.expedientes.application.audit.service import AuditService


@dataclass
class _Audit:
    """Dataclass fake of AuditLogPort — no MagicMock (Hard Rule 5)."""

    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


@dataclass
class _RaisingAudit:
    """Dependency-failure fake: ``append`` raises before any state mutation."""

    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        raise RuntimeError("audit sink unavailable")

    async def record_change(self, change: object) -> None:
        pass


def _make_service(audit: _Audit | None = None) -> tuple[AuditService, _Audit]:
    audit = audit or _Audit()
    return AuditService(audit_log=audit), audit


def _cmd(
    *,
    actor_id: UUID | None,
    target_id: UUID | None = None,
    correlation_id: UUID | None = None,
    payload: dict[str, Any] | None = None,
) -> AuditCommand:
    return AuditCommand(actor_id, target_id or uuid4(), correlation_id, payload or {})


@pytest.mark.parametrize(
    ("method_name", "category"),
    [
        ("record_access", "access"),
        ("record_change", "change"),
        ("record_export", "export"),
        ("record_integration", "integration"),
    ],
)
async def test_record_emits_one_event_with_correct_metadata(
    method_name: str, category: str
) -> None:
    actor, target = uuid4(), uuid4()
    service, audit = _make_service()
    method = getattr(service, method_name)
    result = await method(_cmd(actor_id=actor, target_id=target))
    assert isinstance(result, AuditResult)
    assert result.category == category
    assert result.actor_id == actor
    event = audit.events[0]
    assert event.event_type == f"audit.{category}.recorded"
    assert event.actor_id == actor
    assert event.target_id == target
    assert event.capacidad == "EXP-CAP-046"
    assert event.module == "expedientes"
    assert event.result == "ok"
    assert event.created_at is not None
    assert event.payload["app_id"] == EXPEDIENTES_APP_ID


@pytest.mark.parametrize(
    ("method_name", "actor_id"),
    [
        ("record_access", None),
        ("record_change", None),
        ("record_export", None),
        ("record_integration", None),
        ("record_access", UUID(int=0)),
        ("record_change", UUID(int=0)),
        ("record_export", UUID(int=0)),
        ("record_integration", UUID(int=0)),
    ],
)
async def test_authorization_error_when_actor_invalid(
    method_name: str, actor_id: UUID | None
) -> None:
    service, audit = _make_service()
    with pytest.raises(AuditAuthorizationError, match="actor_id"):
        await getattr(service, method_name)(_cmd(actor_id=actor_id))
    assert audit.events == []


# --- Concurrencia / dependencia -------------------------------------------


async def test_dependency_failure_wraps_runtime_error_and_leaves_no_partial_state() -> None:
    audit = _RaisingAudit()
    service = AuditService(audit_log=audit)
    with pytest.raises(AuditDependencyError) as excinfo:
        await service.record_access(_cmd(actor_id=uuid4()))
    assert isinstance(excinfo.value.__cause__, RuntimeError)
    assert "audit sink unavailable" in str(excinfo.value.__cause__)
    assert audit.events == []


async def test_two_concurrent_services_emit_distinct_events_with_unique_ids() -> None:
    audit = _Audit()
    service_a = AuditService(audit_log=audit)
    service_b = AuditService(audit_log=audit)
    actor_a, actor_b = uuid4(), uuid4()
    await service_a.record_access(_cmd(actor_id=actor_a))
    await service_b.record_export(_cmd(actor_id=actor_b))
    assert len(audit.events) == 2
    assert audit.events[0].id != audit.events[1].id
    assert audit.events[0].actor_id == actor_a
    assert audit.events[1].actor_id == actor_b
    assert audit.events[0].event_type == "audit.access.recorded"
    assert audit.events[1].event_type == "audit.export.recorded"


# --- Invariantes ----------------------------------------------------------


async def test_event_propagates_correlation_id_and_merges_payload_with_app_id() -> None:  # noqa: E501
    service, audit = _make_service()
    correlation = uuid4()
    await service.record_access(
        _cmd(
            actor_id=uuid4(),
            correlation_id=correlation,
            payload={"field": "v0", "row_id": "abc-123"},
        )
    )
    event = audit.events[0]
    assert event.correlation_id == correlation
    assert event.payload["field"] == "v0"
    assert event.payload["row_id"] == "abc-123"
    assert event.payload["app_id"] == EXPEDIENTES_APP_ID


@pytest.mark.parametrize(
    "instance",
    [
        AuditCommand(actor_id=uuid4(), target_id=uuid4(), correlation_id=None, payload={}),
        AuditResult(
            actor_id=uuid4(),
            category="access",
            event_id=uuid4(),
            recorded_at=datetime(2026, 1, 1),
        ),
    ],
)
async def test_audit_dataclass_is_frozen(instance: object) -> None:
    with pytest.raises(dataclasses.FrozenInstanceError):
        instance.actor_id = uuid4()  # type: ignore[attr-defined]


async def test_command_module_does_not_declare_app_id_literal() -> None:
    """EXPEDIENTES_APP_ID = 19 lives ONLY in _evidence.py (DRY ratchet)."""
    from app.src.modules.expedientes.application.audit import command as cmd_mod

    source = inspect.getsource(cmd_mod)
    assert "= 19" not in source, (
        "command.py must not redeclare EXPEDIENTES_APP_ID; the constant "
        "lives in _evidence.py — see BASELINE dup:79d8c5976f0f"
    )
