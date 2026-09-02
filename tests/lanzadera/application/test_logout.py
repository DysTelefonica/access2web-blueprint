# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#540)
"""Tests Categoría 2 (unit / use case) para ``application/logout.py``.

D-W62-2: revocation shortens ``expires_at`` to ``now`` instead of
deleting (DA-11 audit-consumer visibility).
D-W62-4: only the session presented in the request is revoked
("logout everywhere" is a separate epic).
DA-11: every path (success, not-found, double-revoke) emits an
audit row.
HR-2 de la skill ``lanzadera-testing-strategy``: ningún ``MagicMock``;
se usan los fakes de ``tests/lanzadera/_fakes.py``.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from uuid import UUID, uuid4

import pytest

from app.src.modules.lanzadera.application.logout import logout
from app.src.modules.lanzadera.domain.errors import SessionNotFoundError
from app.src.modules.lanzadera.domain.session import Session
from tests.lanzadera._fakes import (
    FakeAuditLog,
    FakeSessionRepository,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _now() -> datetime:
    return datetime(2026, 9, 1, 12, 0, 0, tzinfo=UTC)


async def _seed_session(
    sessions: FakeSessionRepository,
    *,
    user_id: UUID | None = None,
    session_id: UUID | None = None,
    expires_at: datetime | None = None,
) -> Session:
    """Seed a live session (expires_at in the future). NOTE: the
    ``FakeSessionRepository.create`` method is async (DA-1) — must be awaited."""
    sid = session_id or uuid4()
    uid = user_id or uuid4()
    now = _now()
    session = Session(
        id=sid,
        user_id=uid,
        created_at=now - timedelta(hours=1),
        expires_at=expires_at or now + timedelta(hours=1),
    )
    await sessions.create(session)
    return session


# ---------------------------------------------------------------------------
# Success path
# ---------------------------------------------------------------------------


async def test_logout_revokes_session() -> None:
    sessions = FakeSessionRepository()
    audit = FakeAuditLog()
    session = await _seed_session(sessions)
    now = _now()

    await logout(session.id, now=now, sessions=sessions, audit=audit)

    assert sessions.sessions[session.id].expires_at == now


async def test_logout_appends_audit_success_row() -> None:
    sessions = FakeSessionRepository()
    audit = FakeAuditLog()
    session = await _seed_session(sessions)
    now = _now()

    await logout(session.id, now=now, sessions=sessions, audit=audit)

    assert len(audit.entries) == 1
    row = audit.entries[0]
    assert row.event_type == "auth.logout.success"
    assert row.result == "success"
    assert row.target_id == str(session.id)
    assert row.payload["user_id"] == str(session.user_id)


async def test_logout_records_exactly_one_revoke_call() -> None:
    """DA-11: one logout → one revoke. The audit row is NOT counted
    as a revoke (it's a separate side effect on ``audit.append``)."""
    sessions = FakeSessionRepository()
    audit = FakeAuditLog()
    session = await _seed_session(sessions)

    await logout(session.id, now=_now(), sessions=sessions, audit=audit)

    assert len(sessions.revoke_calls) == 1
    assert sessions.revoke_calls[0] == session.id


# ---------------------------------------------------------------------------
# Failure path — unknown session
# ---------------------------------------------------------------------------


async def test_logout_raises_session_not_found_for_unknown_session_id() -> None:
    sessions = FakeSessionRepository()
    audit = FakeAuditLog()
    unknown_id = uuid4()

    with pytest.raises(SessionNotFoundError):
        await logout(unknown_id, now=_now(), sessions=sessions, audit=audit)

    assert sessions.revoke_calls == []


async def test_logout_appends_audit_failure_row_for_unknown_session() -> None:
    sessions = FakeSessionRepository()
    audit = FakeAuditLog()
    unknown_id = uuid4()

    with pytest.raises(SessionNotFoundError):
        await logout(unknown_id, now=_now(), sessions=sessions, audit=audit)

    assert len(audit.entries) == 1
    row = audit.entries[0]
    assert row.event_type == "auth.logout.failure"
    assert row.result == "failure"
    assert row.target_id == str(unknown_id)
    assert row.payload["reason"] == "unknown_session"


# ---------------------------------------------------------------------------
# Idempotency — second revoke succeeds silently
# ---------------------------------------------------------------------------


async def test_logout_is_idempotent_on_second_call() -> None:
    """D-W62-2: revoking an already-revoked session succeeds without
    raising. ``expires_at`` stays at the earlier timestamp (D-W62-2
    + DA-11 audit-consumer visibility)."""
    sessions = FakeSessionRepository()
    audit = FakeAuditLog()
    session = await _seed_session(sessions)
    first = _now()
    second = first + timedelta(minutes=5)

    await logout(session.id, now=first, sessions=sessions, audit=audit)
    await logout(session.id, now=second, sessions=sessions, audit=audit)

    # expires_at stays at the earlier timestamp.
    assert sessions.sessions[session.id].expires_at == first
    # Two audit rows: success then success (idempotent path).
    assert len(audit.entries) == 2
    assert all(e.event_type == "auth.logout.success" for e in audit.entries)
    # Two revoke calls recorded.
    assert sessions.revoke_calls == [session.id, session.id]


# ---------------------------------------------------------------------------
# Input validation
# ---------------------------------------------------------------------------


async def test_logout_raises_value_error_for_none_session_id() -> None:
    sessions = FakeSessionRepository()
    audit = FakeAuditLog()

    with pytest.raises(ValueError):
        await logout(None, now=_now(), sessions=sessions, audit=audit)  # type: ignore[arg-type]
