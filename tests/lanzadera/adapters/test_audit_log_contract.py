"""Contract-conformance test for :class:`AuditLogPort`.

Async I/O uses ``asyncio.run`` (the project's ``pyproject.toml`` does
not enable ``pytest-asyncio``).
"""

from __future__ import annotations

import asyncio
from datetime import UTC, datetime, timedelta
from uuid import uuid4

from app.src.modules.lanzadera.adapters.repos import AuditLogPg
from app.src.modules.lanzadera.domain.audit_event import AuditEvent
from app.src.modules.lanzadera.domain.ports.audit_log import AuditLogPort
from tests.lanzadera.adapters._fakes import FakeAuditLogPort


def test_audit_log_pg_satisfies_protocol() -> None:
    adapter: AuditLogPort = AuditLogPg()
    for method in ("append", "list_for_actor"):
        assert hasattr(adapter, method), f"missing {method}"


def _make_event(event_type: str, actor_id, at: datetime) -> AuditEvent:
    return AuditEvent(
        id=uuid4(),
        event_type=event_type,
        actor_id=actor_id,
        target_id="some-target",
        module="lanzadera",
        result="ok",
        correlation_id=uuid4(),
        payload={"k": "v"},
        created_at=at,
    )


def test_fake_append_persists_event() -> None:
    fake = FakeAuditLogPort()
    event = _make_event("auth.login.success", actor_id=uuid4(), at=datetime.now(UTC))

    async def _go() -> None:
        await fake.append(event)
        assert event in fake.events

    asyncio.run(_go())


def test_fake_list_for_actor_filters_by_actor_and_since() -> None:
    """The list API must respect both filters (D55 — no telemetry leaking)."""
    fake = FakeAuditLogPort()
    actor = uuid4()
    other_actor = uuid4()
    since = datetime.now(UTC)
    later = since + timedelta(minutes=5)

    async def _go() -> None:
        await fake.append(_make_event("auth.login.success", actor, at=since))
        await fake.append(_make_event("auth.login.failure", actor, at=later))
        await fake.append(_make_event("auth.login.success", other_actor, at=since))

        listed = list(await fake.list_for_actor(actor, since=since))
        assert len(listed) == 2
        assert {e.event_type for e in listed} == {"auth.login.success", "auth.login.failure"}

    asyncio.run(_go())


def test_fake_list_for_actor_excludes_events_before_since() -> None:
    fake = FakeAuditLogPort()
    actor = uuid4()
    now = datetime.now(UTC)
    boundary = now + timedelta(seconds=1)

    async def _go() -> None:
        await fake.append(_make_event("auth.login.success", actor, at=now))
        later = _make_event("auth.bootstrap.set_password", actor, at=now + timedelta(minutes=10))
        await fake.append(later)

        listed = list(await fake.list_for_actor(actor, since=boundary))
        assert len(listed) == 1
        assert listed[0].event_type == "auth.bootstrap.set_password"

    asyncio.run(_go())


def test_audit_event_rejects_empty_event_type() -> None:
    """DA-11 + D55: the domain rejects malformed events at the boundary."""
    with pytest.raises(ValueError, match="event_type"):  # noqa: F821
        AuditEvent(
            id=uuid4(),
            event_type="",
            actor_id=None,
            target_id="x",
            module="lanzadera",
            result="ok",
            correlation_id=uuid4(),
            payload={},
            created_at=datetime.now(UTC),
        )


import pytest  # noqa: E402  — after the test that uses it


def test_pg_skeleton_raises_not_implemented() -> None:
    adapter = AuditLogPg()
    event = _make_event("x", actor_id=None, at=datetime.now(UTC))

    async def _go() -> None:
        await adapter.append(event)

    try:
        asyncio.run(_go())
    except NotImplementedError:
        return
    raise AssertionError("expected NotImplementedError")
