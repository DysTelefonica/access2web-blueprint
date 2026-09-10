# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp
"""Tests for ``revoke_assignment`` (issue #588)."""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import UUID, uuid4

import pytest

from app.src.modules.lanzadera.application.revoke_assignment import revoke_assignment
from app.src.modules.lanzadera.domain.assignment import Assignment


class FakeAssignmentRepo:
    """Minimal in-memory assignment repo for the revoke tests."""

    def __init__(self) -> None:
        self.by_id: dict[UUID, Assignment] = {}
        self.revoke_calls: list[tuple[UUID, int]] = []

    async def revoke(self, user_id: UUID, app_id: int, *, now: datetime) -> Assignment | None:
        self.revoke_calls.append((user_id, app_id))
        for a in self.by_id.values():
            if a.user_id == user_id and a.app_id == app_id and a.revoked_at is None:
                a.revoked_at = now
                return a
        return None


class FakeAuditLog:
    def __init__(self) -> None:
        self.entries: list = []

    async def append(self, entry: object) -> None:
        self.entries.append(entry)


@pytest.mark.asyncio
async def test_revoke_sets_revoked_at() -> None:
    """The assignment's ``revoked_at`` is set and the row is returned."""
    assignment_id = uuid4()
    user_id = uuid4()
    app_id = 3
    now = datetime(2026, 9, 10, tzinfo=UTC)

    repo = FakeAssignmentRepo()
    repo.by_id[assignment_id] = Assignment(
        id=assignment_id,
        user_id=user_id,
        app_id=app_id,
        profile_id=uuid4(),
        granted_by=None,
        granted_at=datetime(2026, 1, 1, tzinfo=UTC),
        revoked_at=None,
    )
    audit = FakeAuditLog()

    await revoke_assignment(
        user_id=user_id,
        app_id=app_id,
        assignments=repo,
        audit=audit,
        now=now,
    )

    assignment = repo.by_id[assignment_id]
    assert assignment.revoked_at == now
    assert repo.revoke_calls == [(user_id, app_id)]


@pytest.mark.asyncio
async def test_revoke_raises_when_no_live_assignment() -> None:
    """``ValueError`` when no live assignment exists for the (user, app) pair."""
    user_id = uuid4()
    repo = FakeAssignmentRepo()
    audit = FakeAuditLog()
    now = datetime(2026, 9, 10, tzinfo=UTC)

    with pytest.raises(ValueError) as exc_info:
        await revoke_assignment(
            user_id=user_id,
            app_id=99,
            assignments=repo,
            audit=audit,
            now=now,
        )

    assert "no live assignment" in str(exc_info.value)


@pytest.mark.asyncio
async def test_revoke_emits_audit_event() -> None:
    """A success audit entry is appended after the revocation."""
    assignment_id = uuid4()
    user_id = uuid4()
    app_id = 5
    now = datetime(2026, 9, 10, tzinfo=UTC)
    actor_id = uuid4()

    repo = FakeAssignmentRepo()
    repo.by_id[assignment_id] = Assignment(
        id=assignment_id,
        user_id=user_id,
        app_id=app_id,
        profile_id=uuid4(),
        granted_by=None,
        granted_at=datetime(2026, 1, 1, tzinfo=UTC),
        revoked_at=None,
    )
    audit = FakeAuditLog()

    await revoke_assignment(
        user_id=user_id,
        app_id=app_id,
        assignments=repo,
        audit=audit,
        now=now,
        actor_id=actor_id,
    )

    assert len(audit.entries) == 1
    entry = audit.entries[0]
    assert entry.event_type == "assignments.revoke"
    assert entry.actor_id == actor_id
    assert entry.result == "success"
    assert entry.payload["user_id"] == str(user_id)
    assert entry.payload["app_id"] == app_id


@pytest.mark.asyncio
async def test_revoke_no_audit_on_not_found() -> None:
    """No audit entry when no live assignment exists (the exception propagates)."""
    repo = FakeAssignmentRepo()
    audit = FakeAuditLog()
    now = datetime(2026, 9, 10, tzinfo=UTC)

    with pytest.raises(ValueError):
        await revoke_assignment(
            user_id=uuid4(),
            app_id=99,
            assignments=repo,
            audit=audit,
            now=now,
        )

    assert len(audit.entries) == 0
