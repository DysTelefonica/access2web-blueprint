# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-7, D52, D58 — `App` is the catalog row. Topology and registration status
# are StrEnums that round-trip to the Postgres ENUM wire values.
"""Strict TDD — `App` entity + `AppTopology` + `AppRegistrationStatus` (Phase 1, tasks 1.2)."""

from __future__ import annotations

from datetime import datetime, timezone

import pytest

from app.src.modules.lanzadera.domain.app import App, AppRegistrationStatus, AppTopology

# ---------------------------------------------------------------------------
# Helpers / fixtures
# ---------------------------------------------------------------------------


def _now() -> datetime:
    return datetime(2026, 8, 9, 12, 0, 0, tzinfo=timezone.utc)


def _new_app(
    *,
    app_id: int = 1,
    name: str = "Expedientes",
    short_code: str = "EXP",
    topology: AppTopology = AppTopology.CENTRAL,
    requires_office_presence: bool = False,
    registration: AppRegistrationStatus = AppRegistrationStatus.ACTIVE,
) -> App:
    return App(
        id=app_id,
        name=name,
        short_code=short_code,
        deployment_topology=topology,
        requires_office_presence=requires_office_presence,
        registration_status=registration,
        created_at=_now(),
        updated_at=_now(),
    )


# ---------------------------------------------------------------------------
# Enum round-trips
# ---------------------------------------------------------------------------


class TestAppTopologyStrEnum:
    """`AppTopology` is `StrEnum`; values match the Postgres ENUM."""

    def test_central_member_is_str(self) -> None:
        assert isinstance(AppTopology.CENTRAL, str)

    def test_office_nas_member_is_str(self) -> None:
        assert isinstance(AppTopology.OFFICE_NAS, str)

    @pytest.mark.parametrize(
        ("wire", "member"),
        [
            ("central", AppTopology.CENTRAL),
            ("office-nas", AppTopology.OFFICE_NAS),
        ],
    )
    def test_wire_value_round_trips(self, wire: str, member: AppTopology) -> None:
        assert AppTopology(wire) is member


class TestAppRegistrationStatusStrEnum:
    """`AppRegistrationStatus` is `StrEnum`; values match the Postgres ENUM."""

    def test_pending_member_is_str(self) -> None:
        assert isinstance(AppRegistrationStatus.PENDING, str)

    def test_active_member_is_str(self) -> None:
        assert isinstance(AppRegistrationStatus.ACTIVE, str)

    def test_retired_member_is_str(self) -> None:
        assert isinstance(AppRegistrationStatus.RETIRED, str)

    @pytest.mark.parametrize(
        ("wire", "member"),
        [
            ("pending", AppRegistrationStatus.PENDING),
            ("active", AppRegistrationStatus.ACTIVE),
            ("retired", AppRegistrationStatus.RETIRED),
        ],
    )
    def test_wire_value_round_trips(self, wire: str, member: AppRegistrationStatus) -> None:
        assert AppRegistrationStatus(wire) is member


# ---------------------------------------------------------------------------
# Construction — happy path
# ---------------------------------------------------------------------------


class TestAppConstruction:
    def test_app_carries_catalog_attributes(self) -> None:
        app = _new_app()
        assert app.id == 1
        assert app.name == "Expedientes"
        assert app.short_code == "EXP"
        assert app.deployment_topology is AppTopology.CENTRAL
        assert app.requires_office_presence is False
        assert app.registration_status is AppRegistrationStatus.ACTIVE

    def test_central_app_without_office_requirement(self) -> None:
        """DA-7 / 0002: `EjecucionEnOficina = 'No'` maps to `central` + False."""
        app = _new_app(topology=AppTopology.CENTRAL, requires_office_presence=False)
        assert app.deployment_topology is AppTopology.CENTRAL
        assert app.requires_office_presence is False

    def test_office_nas_app_with_office_requirement(self) -> None:
        """DA-7 / 0002: `EjecucionEnOficina = 'Sí'` maps to `office-nas` + True."""
        app = _new_app(topology=AppTopology.OFFICE_NAS, requires_office_presence=True)
        assert app.deployment_topology is AppTopology.OFFICE_NAS
        assert app.requires_office_presence is True

    def test_short_code_must_be_non_empty(self) -> None:
        with pytest.raises(ValueError, match="short_code"):
            _new_app(short_code="")

    def test_name_must_be_non_empty(self) -> None:
        with pytest.raises(ValueError, match="name"):
            _new_app(name="")


# ---------------------------------------------------------------------------
# Lifecycle transitions
# ---------------------------------------------------------------------------


class TestAppLifecycleTransitions:
    """Registration transitions: pending -> active -> retired (terminal)."""

    def test_pending_to_active(self) -> None:
        app = _new_app(registration=AppRegistrationStatus.PENDING)
        app.registration_status = AppRegistrationStatus.ACTIVE
        assert app.registration_status is AppRegistrationStatus.ACTIVE

    def test_active_to_retired(self) -> None:
        app = _new_app(registration=AppRegistrationStatus.ACTIVE)
        app.registration_status = AppRegistrationStatus.RETIRED
        assert app.registration_status is AppRegistrationStatus.RETIRED

    def test_app_is_mutable(self) -> None:
        """Catalog rows evolve (status, name, updated_at) -> not frozen."""
        app = _new_app()
        app.updated_at = _now()
        app.name = "Expedientes v2"
        assert app.name == "Expedientes v2"
