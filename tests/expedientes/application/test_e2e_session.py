"""Strict TDD for EXP-CAP-041 E2E session (E07)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.e2e_session.command import (
    E2ESessionAuthorizationError,
    E2ESessionError,
    E2ESessionStatus,
    E2ESessionValidationError,
)
from app.src.modules.expedientes.application.e2e_session.service import (
    E2ESessionService,
)


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


class _SessionStore:
    """In-memory session store used by the test."""

    def __init__(self) -> None:
        self.by_id: dict[UUID, dict[str, Any]] = {}
        self.fail_on: set[UUID] = set()

    async def create(self, session: dict[str, Any]) -> None:
        if session["session_id"] in self.fail_on:
            raise RuntimeError("store failed on create")
        self.by_id[session["session_id"]] = session

    async def get(self, session_id: UUID) -> dict[str, Any] | None:
        return self.by_id.get(session_id)

    async def update(self, session: dict[str, Any]) -> None:
        if session["session_id"] in self.fail_on:
            raise RuntimeError("store failed on update")
        self.by_id[session["session_id"]] = session


def _make_service() -> tuple[E2ESessionService, _Audit, _SessionStore]:
    audit = _Audit()
    store = _SessionStore()
    service = E2ESessionService(
        audit_log=audit,
        store=store,
        permissions=set(),
    )
    return service, audit, store


async def test_open_creates_active_session() -> None:
    service, audit, _ = _make_service()
    service.grant("e2e.session")
    actor = uuid4()

    result = await service.open(actor_id=actor)

    assert result.status is E2ESessionStatus.ACTIVE
    assert result.actor_id == actor
    assert result.session_id is not None
    assert result.opened_at <= result.last_seen_at
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "e2e.session.opened"
    assert audit.events[0].capacidad == "EXP-CAP-041"


async def test_resume_returns_existing_session() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.session")
    actor = uuid4()
    first = await service.open(actor_id=actor)

    second = await service.resume(actor_id=actor, session_id=first.session_id)

    assert second.session_id == first.session_id
    assert second.status is E2ESessionStatus.ACTIVE
    assert second.last_seen_at >= first.last_seen_at


async def test_resume_rejects_other_users_session() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.session")
    owner = uuid4()
    intruder = uuid4()
    opened = await service.open(actor_id=owner)

    with pytest.raises(E2ESessionAuthorizationError, match="does not belong"):
        await service.resume(actor_id=intruder, session_id=opened.session_id)


async def test_close_marks_session_closed() -> None:
    service, audit, store = _make_service()
    service.grant("e2e.session")
    actor = uuid4()
    opened = await service.open(actor_id=actor)

    closed = await service.close(actor_id=actor, session_id=opened.session_id)

    assert closed.status is E2ESessionStatus.CLOSED
    assert store.by_id[opened.session_id]["status"] == E2ESessionStatus.CLOSED.value
    # 3 events: open (audit), closed (audit), change record (audit).
    assert len(audit.events) == 3
    assert audit.events[1].event_type == "e2e.session.closed"


async def test_resume_on_closed_session_rejected() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.session")
    actor = uuid4()
    opened = await service.open(actor_id=actor)
    await service.close(actor_id=actor, session_id=opened.session_id)

    with pytest.raises(E2ESessionValidationError, match="closed"):
        await service.resume(actor_id=actor, session_id=opened.session_id)


async def test_resume_unknown_session_rejected() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.session")
    with pytest.raises(E2ESessionValidationError, match="not found"):
        await service.resume(actor_id=uuid4(), session_id=uuid4())


async def test_missing_actor_rejected() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.session")
    with pytest.raises(E2ESessionAuthorizationError):
        await service.open(actor_id=None)


async def test_missing_permission_rejected() -> None:
    service, _, _ = _make_service()
    with pytest.raises(E2ESessionAuthorizationError, match="e2e.session"):
        await service.open(actor_id=uuid4())


async def test_dependency_failure_rolls_back() -> None:
    service, audit, store = _make_service()
    service.grant("e2e.session")

    # Pre-seed the store to inject a failure at create time.
    class _FailingStore(_SessionStore):
        async def create(self, session: dict[str, Any]) -> None:
            raise RuntimeError("store unavailable")

    service._store = _FailingStore()

    with pytest.raises(E2ESessionError, match="session failed"):
        await service.open(actor_id=uuid4())

    assert not audit.events


async def test_open_actor_does_not_match_audit_event_actor() -> None:
    service, audit, _ = _make_service()
    service.grant("e2e.session")
    actor = uuid4()
    result = await service.open(actor_id=actor)
    assert audit.events[0].actor_id == actor
    assert result.actor_id == actor
