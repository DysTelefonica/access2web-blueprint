"""Contract-conformance test for :class:`AuditLogPort`.

Pin both implementations against the same Protocol:
- the in-memory ``FakeAuditLogPort`` (used by the rest of the test
  suite), and
- the Postgres ``AuditLogPg`` wired to the
  :func:`async_session_factory` seam (DA-1, W03 #55).
"""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import uuid4

from app.src.modules.lanzadera.adapters.persistence import (
    AuditLogPg,
    async_session_factory,
)
from app.src.modules.lanzadera.domain.audit_event import AuditEvent
from app.src.modules.lanzadera.domain.ports.audit_log import AuditLogPort
from tests.lanzadera.adapters._fakes import FakeAuditLogPort


def test_audit_log_pg_satisfies_protocol() -> None:
    """Static structural check on the Postgres adapter's method set."""
    expected = {"append", "list_for_actor"}
    actual = set(dir(AuditLogPg))
    missing = expected - actual
    assert not missing, f"AuditLogPg is missing Protocol methods: {missing}"


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


async def test_fake_append_persists_event() -> None:
    fake: AuditLogPort = FakeAuditLogPort()
    event = _make_event("auth.login.success", actor_id=uuid4(), at=datetime.now(UTC))

    await fake.append(event)
    assert event in fake.events


async def test_audit_log_pg_construction() -> None:
    """``AuditLogPg`` takes the factory seam and is constructable (DA-1).

    We do not connect to Postgres here; an unreachable URL is fine
    because the engine is created lazily. The test merely verifies
    the constructor's wiring shape.
    """
    _, factory = async_session_factory("postgresql+asyncpg://nobody:nopwd@127.0.0.1:1/nodb")
    adapter = AuditLogPg(factory)
    assert adapter is not None
