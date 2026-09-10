# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Tests Categoría 2 (unit / use case) para ``application/get_my_apps.py``.

D22 + DA-12 + H11: assignments son la única fuente de verdad para permisos.
Los tests verifican el join assignments → apps → profiles → capabilities.

HR-2 de la skill ``lanzadera-testing-strategy``: ningún ``MagicMock``;
se usan los fakes de ``tests/lanzadera/_fakes.py``.
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import TYPE_CHECKING
from uuid import uuid4

import pytest

from app.src.modules.lanzadera.application.get_my_apps import (
    EffectiveApp,
    get_my_apps,
    get_my_capabilities_for_app,
)
from app.src.modules.lanzadera.domain.app import App, AppRegistrationStatus, AppTopology
from app.src.modules.lanzadera.domain.assignment import Assignment
from app.src.modules.lanzadera.domain.profile import Profile
from tests.lanzadera._fakes import (
    FakeAppRepository,
    FakeAssignmentRepository,
    FakeProfileRepository,
)

if TYPE_CHECKING:
    from uuid import UUID


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _utc(year: int, month: int, day: int) -> datetime:
    return datetime(year, month, day, tzinfo=UTC)


# ---------------------------------------------------------------------------
# get_my_apps
# ---------------------------------------------------------------------------


class TestGetMyApps:
    @pytest.fixture
    def app_repo(self) -> FakeAppRepository:
        return FakeAppRepository()

    @pytest.fixture
    def profile_repo(self) -> FakeProfileRepository:
        return FakeProfileRepository()

    @pytest.fixture
    def assignment_repo(
        self, profile_repo: FakeProfileRepository
    ) -> FakeAssignmentRepository:
        return FakeAssignmentRepository().with_profiles(profile_repo)

    @pytest.mark.asyncio
    async def test_empty_when_no_assignments(
        self,
        assignment_repo: FakeAssignmentRepository,
        app_repo: FakeAppRepository,
        profile_repo: FakeProfileRepository,
    ) -> None:
        user_id = uuid4()
        result = await get_my_apps(
            user_id,
            assignments=assignment_repo,
            apps=app_repo,
            profiles=profile_repo,
        )
        assert result == ()

    @pytest.mark.asyncio
    async def test_returns_app_with_profile_and_capabilities(
        self,
        assignment_repo: FakeAssignmentRepository,
        app_repo: FakeAppRepository,
        profile_repo: FakeProfileRepository,
    ) -> None:
        user_id = uuid4()
        app = App(
            id=1,
            name="Expedientes",
            short_code="EXP",
            deployment_topology=AppTopology.CENTRAL,
            requires_office_presence=False,
            registration_status=AppRegistrationStatus.ACTIVE,
            created_at=_utc(2026, 1, 1),
            updated_at=_utc(2026, 1, 1),
        )
        app_repo.add(app)

        profile = Profile(
            id=uuid4(),
            app_id=1,
            code="ADMIN",
            name="Administrador",
            capabilities={"Calidad": True, "read": True},
            active=True,
            created_at=_utc(2026, 1, 1),
            updated_at=_utc(2026, 1, 1),
        )
        profile_repo.add(profile)

        assignment = Assignment(
            id=uuid4(),
            user_id=user_id,
            app_id=1,
            profile_id=profile.id,
            granted_by=None,
            granted_at=_utc(2026, 1, 1),
            revoked_at=None,
        )
        assignment_repo.add(assignment)

        result = await get_my_apps(
            user_id,
            assignments=assignment_repo,
            apps=app_repo,
            profiles=profile_repo,
        )

        assert len(result) == 1
        assert result[0] == EffectiveApp(
            app_id=1,
            app_name="Expedientes",
            app_short_code="EXP",
            profile_code="ADMIN",
            profile_name="Administrador",
            capabilities=("Calidad", "read"),
        )

    @pytest.mark.asyncio
    async def test_skips_app_not_in_repo(
        self,
        assignment_repo: FakeAssignmentRepository,
        app_repo: FakeAppRepository,
        profile_repo: FakeProfileRepository,
    ) -> None:
        """App not added to the repository is silently dropped from results."""
        user_id = uuid4()
        profile = Profile(
            id=uuid4(),
            app_id=99,
            code="DEFAULT",
            name="Default",
            capabilities={},
            active=True,
            created_at=_utc(2026, 1, 1),
            updated_at=_utc(2026, 1, 1),
        )
        profile_repo.add(profile)

        # Assignment for an app that doesn't exist in the catalog.
        assignment_repo.add(
            Assignment(
                id=uuid4(),
                user_id=user_id,
                app_id=99,
                profile_id=profile.id,
                granted_by=None,
                granted_at=_utc(2026, 1, 1),
                revoked_at=None,
            )
        )

        result = await get_my_apps(
            user_id,
            assignments=assignment_repo,
            apps=app_repo,
            profiles=profile_repo,
        )
        # app_repo.get_by_id(99) returns None -> app is silently dropped.
        assert len(result) == 0


class TestGetMyCapabilitiesForApp:
    @pytest.fixture
    def assignment_repo(
        self,
    ) -> FakeAssignmentRepository:
        return FakeAssignmentRepository()

    @pytest.mark.asyncio
    async def test_empty_when_no_assignment(
        self,
        assignment_repo: FakeAssignmentRepository,
    ) -> None:
        user_id = uuid4()
        result = await get_my_capabilities_for_app(
            user_id,
            app_id=1,
            assignments=assignment_repo,
        )
        assert result == []

    @pytest.mark.asyncio
    async def test_returns_capabilities_for_app(
        self,
        assignment_repo: FakeAssignmentRepository,
    ) -> None:
        profile_repo = FakeProfileRepository()
        assignment_repo = FakeAssignmentRepository().with_profiles(profile_repo)

        user_id = uuid4()
        profile = Profile(
            id=uuid4(),
            app_id=5,
            code="CALIDAD",
            name="Calidad",
            capabilities={"Calidad": True, "write": True},
            active=True,
            created_at=_utc(2026, 1, 1),
            updated_at=_utc(2026, 1, 1),
        )
        profile_repo.add(profile)

        assignment_repo.add(
            Assignment(
                id=uuid4(),
                user_id=user_id,
                app_id=5,
                profile_id=profile.id,
                granted_by=None,
                granted_at=_utc(2026, 1, 1),
                revoked_at=None,
            )
        )

        result = await get_my_capabilities_for_app(
            user_id,
            app_id=5,
            assignments=assignment_repo,
        )
        assert set(result) == {"Calidad", "write"}
