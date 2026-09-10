# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp
"""Tests for ``seed_profiles`` (issue #587).

Verifies:
- All 8 apps × 8 profile codes = 64 rows are created on first run.
- Re-running is idempotent (0 rows created the second time).
- ``SIN_ACCESO`` is not skipped.
"""

from __future__ import annotations

from datetime import UTC, datetime

import pytest

from app.src.modules.lanzadera.application.seed_profiles import (
    _APP_CATALOGUE,
    _PROFILE_CODES,
    seed_profiles,
)
from app.src.modules.lanzadera.domain.profile import Profile


class FakeProfileRepository:
    """In-memory profile repository that records calls and returns None for lookups."""

    def __init__(self) -> None:
        self.by_key: dict[tuple[int, str], Profile] = {}
        self.create_calls: list[Profile] = []

    async def get_by_code(self, app_id: int, code: str) -> Profile | None:
        return self.by_key.get((app_id, code))

    async def create(self, profile: Profile) -> None:
        self.by_key[(profile.app_id, profile.code)] = profile
        self.create_calls.append(profile)


class FakeAuditLog:
    def __init__(self) -> None:
        self.entries: list = []

    async def append(self, entry: object) -> None:
        self.entries.append(entry)


@pytest.mark.asyncio
async def test_first_run_creates_all_rows() -> None:
    """All 8 apps × 8 profile codes are created on the first call."""
    repo = FakeProfileRepository()
    audit = FakeAuditLog()
    now = datetime(2026, 9, 10, tzinfo=UTC)

    created = await seed_profiles(profiles=repo, audit=audit, now=now)

    assert created == len(_APP_CATALOGUE) * len(_PROFILE_CODES)
    assert len(repo.create_calls) == created

    # Verify every (app, code) pair was created
    for app in _APP_CATALOGUE:
        for pdef in _PROFILE_CODES:
            profile = await repo.get_by_code(app["id"], pdef["code"])
            assert profile is not None, f"Missing {app['id']}/{pdef['code']}"
            assert profile.name == pdef["name"]
            assert profile.capabilities == {}  # intentionally empty
            assert profile.active is True
            assert profile.created_at == now


@pytest.mark.asyncio
async def test_second_run_is_idempotent() -> None:
    """Re-running seed_profiles creates no new rows."""
    repo = FakeProfileRepository()
    audit = FakeAuditLog()
    now = datetime(2026, 9, 10, tzinfo=UTC)

    # First run
    first = await seed_profiles(profiles=repo, audit=audit, now=now)
    assert first == len(_APP_CATALOGUE) * len(_PROFILE_CODES)

    # Second run — nothing new
    second = await seed_profiles(profiles=repo, audit=audit, now=now)
    assert second == 0
    assert len(repo.create_calls) == first


@pytest.mark.asyncio
async def test_audit_logged_on_creation() -> None:
    """An audit entry is appended when at least one row is created."""
    repo = FakeProfileRepository()
    audit = FakeAuditLog()

    await seed_profiles(profiles=repo, audit=audit)

    assert len(audit.entries) == 1
    entry = audit.entries[0]
    assert entry.event_type == "profiles.seed"
    assert entry.result == "success"
    assert entry.payload["created"] == len(_APP_CATALOGUE) * len(_PROFILE_CODES)


@pytest.mark.asyncio
async def test_audit_not_logged_when_nothing_created() -> None:
    """No audit entry is appended on a no-op re-run."""
    repo = FakeProfileRepository()
    audit = FakeAuditLog()
    now = datetime(2026, 9, 10, tzinfo=UTC)

    await seed_profiles(profiles=repo, audit=audit, now=now)
    await seed_profiles(profiles=repo, audit=audit, now=now)

    # One entry from first run (second run created nothing)
    assert len(audit.entries) == 1


@pytest.mark.asyncio
async def test_sin_acceso_is_seeded() -> None:
    """SIN_ACCESO profile is not skipped by the exclusivity rule."""
    repo = FakeProfileRepository()
    audit = FakeAuditLog()

    await seed_profiles(profiles=repo, audit=audit)

    for app in _APP_CATALOGUE:
        profile = await repo.get_by_code(app["id"], "SIN_ACCESO")
        assert profile is not None, f"SIN_ACCESO missing for app {app['id']}"
