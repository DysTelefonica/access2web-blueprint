# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-12, H11 — `Assignment` is the `(user, app, profile)` triple. Revocation
# is a soft-delete (`revoked_at`), keeping the row visible to historical reads.
"""Strict TDD — `Assignment` entity (Phase 1, task 1.4)."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from uuid import uuid4

import pytest

from app.src.modules.lanzadera.domain.assignment import Assignment

# ---------------------------------------------------------------------------
# Helpers / fixtures
# ---------------------------------------------------------------------------


def _now() -> datetime:
    return datetime(2026, 8, 9, 12, 0, 0, tzinfo=timezone.utc)


def _new_assignment(
    *,
    granted_by: object = None,
    revoked_at: datetime | None = None,
    granted_at: datetime | None = None,
) -> Assignment:
    return Assignment(
        id=uuid4(),
        user_id=uuid4(),
        app_id=1,
        profile_id=uuid4(),
        granted_by=granted_by,  # type: ignore[arg-type]
        granted_at=granted_at or _now(),
        revoked_at=revoked_at,
    )


# ---------------------------------------------------------------------------
# Construction
# ---------------------------------------------------------------------------


class TestAssignmentConstruction:
    def test_assignment_carries_required_attributes(self) -> None:
        user_id = uuid4()
        profile_id = uuid4()
        grant_id = uuid4()
        assignment = Assignment(
            id=grant_id,
            user_id=user_id,
            app_id=3,
            profile_id=profile_id,
            granted_by=None,
            granted_at=_now(),
            revoked_at=None,
        )
        assert assignment.id == grant_id
        assert assignment.user_id == user_id
        assert assignment.app_id == 3
        assert assignment.profile_id == profile_id
        assert assignment.granted_by is None
        assert assignment.revoked_at is None

    def test_assignment_with_granted_by_admin(self) -> None:
        admin_id = uuid4()
        assignment = _new_assignment(granted_by=admin_id)
        assert assignment.granted_by == admin_id

    def test_granted_by_can_be_none(self) -> None:
        """Self-assignments (seed migrations, bootstrap) carry no actor."""
        assignment = _new_assignment(granted_by=None)
        assert assignment.granted_by is None


# ---------------------------------------------------------------------------
# Revocation (soft delete)
# ---------------------------------------------------------------------------


class TestAssignmentRevocation:
    def test_active_assignment_has_no_revoked_at(self) -> None:
        assignment = _new_assignment()
        assert assignment.revoked_at is None

    def test_revocation_sets_revoked_at(self) -> None:
        """DA-12: revocation is a soft-delete; the row stays readable for audit."""
        assignment = _new_assignment()
        revoked_at = _now() + timedelta(hours=1)
        assignment.revoked_at = revoked_at
        assert assignment.revoked_at == revoked_at

    def test_reactivation_after_revocation_keeps_history(self) -> None:
        """Soft-delete semantics: revoked rows stay revoked; a new grant is a new row."""
        assignment = _new_assignment()
        assignment.revoked_at = _now() + timedelta(hours=1)
        assert assignment.revoked_at is not None
