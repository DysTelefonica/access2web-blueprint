# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# D22, D45, D110, DA-12 — `Profile` carries a JSONB-shaped `capabilities` map.
# G-2 OPEN: capabilities seeded with `{}`; the canonical shape arrives in PR 3a.
"""Strict TDD — `Profile` entity (Phase 1, task 1.3)."""

from __future__ import annotations

from datetime import datetime, timezone
from uuid import uuid4

import pytest

from app.src.modules.lanzadera.domain.profile import Profile

# ---------------------------------------------------------------------------
# Helpers / fixtures
# ---------------------------------------------------------------------------


def _now() -> datetime:
    return datetime(2026, 8, 9, 12, 0, 0, tzinfo=timezone.utc)


def _new_profile(
    *,
    code: str = "default",
    name: str = "Default",
    capabilities: dict | None = None,
    active: bool = True,
    app_id: int = 1,
) -> Profile:
    return Profile(
        id=uuid4(),
        app_id=app_id,
        code=code,
        name=name,
        capabilities=capabilities if capabilities is not None else {},
        active=active,
        created_at=_now(),
        updated_at=_now(),
    )


# ---------------------------------------------------------------------------
# Construction — happy path
# ---------------------------------------------------------------------------


class TestProfileConstruction:
    def test_profile_carries_required_attributes(self) -> None:
        profile_id = uuid4()
        profile = Profile(
            id=profile_id,
            app_id=7,
            code="ADMIN",
            name="Administrator",
            capabilities={"can_edit": True, "max_records": 1000},
            active=True,
            created_at=_now(),
            updated_at=_now(),
        )
        assert profile.id == profile_id
        assert profile.app_id == 7
        assert profile.code == "ADMIN"
        assert profile.name == "Administrator"
        assert profile.capabilities == {"can_edit": True, "max_records": 1000}
        assert profile.active is True

    def test_default_capabilities_is_empty_dict(self) -> None:
        """G-2 OPEN: seed with `{}` until product confirms the canonical shape."""
        profile = _new_profile()
        assert profile.capabilities == {}

    def test_capabilities_accepts_string_number_boolean_values(self) -> None:
        """D45 / D110: capabilities values are `string | number | boolean`."""
        caps: dict = {
            "scope": "global",
            "max_records": 500,
            "can_publish": True,
            "can_delete": False,
        }
        profile = _new_profile(capabilities=caps)
        assert profile.capabilities == caps

    def test_code_must_be_non_empty(self) -> None:
        with pytest.raises(ValueError, match="code"):
            _new_profile(code="")

    def test_name_must_be_non_empty(self) -> None:
        with pytest.raises(ValueError, match="name"):
            _new_profile(name="")


# ---------------------------------------------------------------------------
# Lifecycle — deactivate / reactivate
# ---------------------------------------------------------------------------


class TestProfileLifecycle:
    def test_deactivation_is_allowed(self) -> None:
        """DA-12: deactivated profile remains for historical assignment lookups."""
        profile = _new_profile(active=True)
        profile.active = False
        assert profile.active is False

    def test_reactivation_is_allowed(self) -> None:
        profile = _new_profile(active=False)
        profile.active = True
        assert profile.active is True

    def test_profile_is_mutable(self) -> None:
        """Profile rows evolve (capabilities, name, active) -> not frozen."""
        profile = _new_profile()
        profile.capabilities = {"new_capability": True}
        assert profile.capabilities == {"new_capability": True}
