# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W61
# W61 (#524) — unit tests for the app CRUD use cases.
"""Unit tests for the W61 (#524) app CRUD use cases.

The tests cover the three application-layer entry points the W61 JSON
routes rely on:

- ``create_app`` — POST ``/admin/apps`` path.
- ``update_app`` — PATCH ``/admin/apps/{id}`` path.
- ``disable_app`` — DELETE ``/admin/apps/{id}`` path.

Each test seeds the in-memory ``FakeAppRepository`` directly so the
assertions pin the contract the Postgres adapter has to honour: the
use case is a thin pass-through to the driven port.
"""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import uuid4

from app.src.modules.lanzadera.application.create_app import create_app
from app.src.modules.lanzadera.application.disable_app import disable_app
from app.src.modules.lanzadera.application.update_app import update_app
from app.src.modules.lanzadera.domain.app import (
    App,
    AppRegistrationStatus,
    AppTopology,
)
from tests.lanzadera._fakes import FakeAppRepository


def _seed_app(
    *,
    app_id: int = 1,
    name: str = "Expedientes",
    short_code: str = "EXP",
    topology: AppTopology = AppTopology.CENTRAL,
    requires_office: bool = False,
    status: AppRegistrationStatus = AppRegistrationStatus.ACTIVE,
) -> App:
    """Return a seeded ``App`` for the assertion helpers."""
    now = datetime(2026, 1, 1, tzinfo=UTC)
    return App(
        id=app_id,
        name=name,
        short_code=short_code,
        deployment_topology=topology,
        requires_office_presence=requires_office,
        registration_status=status,
        created_at=now,
        updated_at=now,
    )


# ---------------------------------------------------------------------------
# create_app
# ---------------------------------------------------------------------------


async def test_create_app_returns_persisted_app() -> None:
    """``create_app`` returns the row the repo persisted with the
    server-defaulted ``registration_status='pending'``."""
    apps = FakeAppRepository()

    app = await create_app(
        "Lanzadera",
        short_code="lanza",
        deployment_topology=AppTopology.CENTRAL,
        requires_office_presence=False,
        apps=apps,
    )

    # The fake mints ids off ``next_id`` starting at 100 — the exact
    # value is incidental; what matters is that the returned App
    # matches the row the fake recorded in ``create_calls``.
    assert app.name == "Lanzadera"
    assert app.short_code == "lanza"
    assert app.deployment_topology is AppTopology.CENTRAL
    assert app.requires_office_presence is False
    assert app.registration_status is AppRegistrationStatus.PENDING
    assert app.id == apps.next_id
    assert apps.by_id[app.id] == app
    assert apps.create_calls == [app]


async def test_create_app_records_call_in_fake() -> None:
    """``create_app`` calls the repo exactly once and forwards every kwarg."""
    apps = FakeAppRepository()

    await create_app(
        "Expedientes",
        short_code="EXP",
        deployment_topology=AppTopology.OFFICE_NAS,
        requires_office_presence=True,
        apps=apps,
        actor_id=uuid4(),
    )

    assert len(apps.create_calls) == 1
    persisted = apps.create_calls[0]
    assert persisted.name == "Expedientes"
    assert persisted.short_code == "EXP"
    assert persisted.deployment_topology is AppTopology.OFFICE_NAS
    assert persisted.requires_office_presence is True
    # ``registration_status`` defaults to ``pending`` so non-admin users
    # cannot see the row until an activation flow ships (DA-7).
    assert persisted.registration_status is AppRegistrationStatus.PENDING


async def test_create_app_propagates_validation_error() -> None:
    """``create_app`` propagates the ``ValueError`` raised by ``App.__post_init__``.

    The dataclass rejects empty ``name`` / ``short_code``; the use case
    surfaces the failure without wrapping it (the route layer renders
    the 400 from the propagated error).
    """
    import pytest

    apps = FakeAppRepository()

    with pytest.raises(ValueError, match="non-empty"):
        await create_app(
            "   ",  # whitespace-only name
            short_code="EXP",
            deployment_topology=AppTopology.CENTRAL,
            requires_office_presence=False,
            apps=apps,
        )

    # The repo did NOT receive the call — validation runs before the
    # adapter is reached.
    assert apps.create_calls == []


# ---------------------------------------------------------------------------
# update_app
# ---------------------------------------------------------------------------


async def test_update_app_returns_patched_row() -> None:
    """``update_app`` returns the post-update ``App`` (with a fresh ``updated_at``)."""
    apps = FakeAppRepository()
    seeded = _seed_app(app_id=42, name="Expedientes", short_code="EXP")
    apps.add(seeded)
    original_updated_at = seeded.updated_at

    app = await update_app(
        42,
        apps=apps,
        name="Expedientes v2",
        actor_id=uuid4(),
    )

    assert app.name == "Expedientes v2"
    assert app.short_code == "EXP"  # untouched
    assert app.deployment_topology is AppTopology.CENTRAL  # untouched
    assert app.requires_office_presence is False  # untouched
    assert app.updated_at >= original_updated_at
    # The fake records every update call so the route can pin the
    # exact patch shape it forwarded.
    assert apps.update_calls == [(42, {"name": "Expedientes v2"})]


async def test_update_app_patches_partial_fields() -> None:
    """``update_app`` honours the partial-patch contract (None fields are skipped)."""
    apps = FakeAppRepository()
    apps.add(_seed_app(app_id=7, name="Expedientes", requires_office=False))

    app = await update_app(
        7,
        apps=apps,
        requires_office_presence=True,
        # name and deployment_topology are intentionally absent
    )

    assert app.requires_office_presence is True
    assert app.name == "Expedientes"  # untouched
    assert app.deployment_topology is AppTopology.CENTRAL  # untouched
    assert apps.update_calls == [(7, {"requires_office_presence": True})]


async def test_update_app_noop_with_no_kwargs() -> None:
    """``update_app`` with every kwarg ``None`` is a no-op (only ``updated_at`` moves)."""
    apps = FakeAppRepository()
    original = _seed_app(app_id=11)
    apps.add(original)

    app = await update_app(11, apps=apps)

    # The repo received the call with an empty kwargs payload; the
    # only column change is ``updated_at`` (the fake re-stamps it).
    assert apps.update_calls == [(11, {})]
    assert app.name == original.name
    assert app.short_code == original.short_code
    assert app.updated_at >= original.updated_at


# ---------------------------------------------------------------------------
# disable_app
# ---------------------------------------------------------------------------


async def test_disable_app_returns_retired_row() -> None:
    """``disable_app`` flips ``registration_status`` to ``retired`` and returns the row."""
    apps = FakeAppRepository()
    apps.add(_seed_app(app_id=99, status=AppRegistrationStatus.ACTIVE))

    app = await disable_app(99, apps=apps, actor_id=uuid4())

    assert app.registration_status is AppRegistrationStatus.RETIRED
    assert app.id == 99
    assert app.name == "Expedientes"  # other fields untouched
    # The fake records the disable so the route can pin the exact
    # id the destructive call targeted.
    assert apps.disable_calls == [99]
    # The by_id entry is the post-disable row.
    assert apps.by_id[99] == app


async def test_disable_app_propagates_unknown_id() -> None:
    """``disable_app`` propagates the ``RuntimeError`` when the id is unknown.

    The Postgres adapter raises ``RuntimeError`` when ``RETURNING``
    returns an empty result set (the row was deleted between the
    route's read and write). The use case surfaces the failure
    without wrapping it.
    """
    import pytest

    apps = FakeAppRepository()

    with pytest.raises(RuntimeError, match="no row"):
        await disable_app(999, apps=apps)


__all__ = [
    "test_create_app_propagates_validation_error",
    "test_create_app_records_call_in_fake",
    "test_create_app_returns_persisted_app",
    "test_disable_app_propagates_unknown_id",
    "test_disable_app_returns_retired_row",
    "test_update_app_noop_with_no_kwargs",
    "test_update_app_patches_partial_fields",
    "test_update_app_returns_patched_row",
]
