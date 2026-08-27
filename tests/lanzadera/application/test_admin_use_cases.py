# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Tests for the admin use cases (W54, issue #43)."""

from __future__ import annotations

from datetime import UTC, datetime

import pytest

from app.src.modules.lanzadera.application.assign_profile import assign_profile
from app.src.modules.lanzadera.application.audit_append import audit_append
from app.src.modules.lanzadera.application.bootstrap_global_admins import (
    bootstrap_global_admins,
)
from app.src.modules.lanzadera.application.create_user import create_user
from app.src.modules.lanzadera.application.disable_user import disable_user
from app.src.modules.lanzadera.application.grant_global_admin import (
    grant_global_admin,
)
from app.src.modules.lanzadera.application.list_effective_apps import (
    list_effective_apps,
)
from app.src.modules.lanzadera.application.revoke_global_admin import (
    revoke_global_admin,
)
from app.src.modules.lanzadera.application.set_password import set_password
from app.src.modules.lanzadera.domain.app import App
from app.src.modules.lanzadera.domain.errors import (
    DuplicateEmailError,
    UserNotFoundError,
)
from app.src.modules.lanzadera.domain.profile import Profile
from app.src.modules.lanzadera.domain.user import User, UserStatus
from tests.lanzadera.application._fakes import (
    FakeAppRepository,
    FakeAssignmentRepository,
    FakeCredentialHasher,
    FakeGlobalAdminRepository,
    FakeProfileRepository,
    FakeSecretManager,
    FakeUserRepository,
)
from tests.lanzadera.auth._fakes import FakeAuditLog


def _now() -> datetime:
    return datetime(2026, 1, 1, tzinfo=UTC)


def _user(email: str = "[email protected]", status: UserStatus = UserStatus.ACTIVE) -> User:
    from uuid import uuid4

    return User(
        id=uuid4(),
        email=email,
        name="Test User",
        dni_encrypted=b"\x00",
        password_hash=None,
        status=status,
        failed_attempts=0,
        last_login_at=None,
        created_at=_now(),
        updated_at=_now(),
    )


# ---------------------------------------------------------------------------
# create_user
# ---------------------------------------------------------------------------


async def test_create_user_persists_with_reset_required_status() -> None:
    users = FakeUserRepository()
    audit = FakeAuditLog()
    secrets = FakeSecretManager()
    user = await create_user(
        "[email protected]",
        name="Alice",
        national_id="12345678",
        users=users,
        audit=audit,
        secrets=secrets,
        now=_now(),
    )
    assert user.email == "[email protected]"  # normalised
    assert user.status is UserStatus.PASSWORD_RESET_REQUIRED
    assert user.password_hash is None  # D89 — bootstrap via reset only
    assert len(audit.entries) == 1
    assert audit.entries[0].event_type == "users.create"


async def test_create_user_rejects_duplicate_email() -> None:
    users = FakeUserRepository()
    users.add(_user("[email protected]"))
    audit = FakeAuditLog()
    secrets = FakeSecretManager()
    with pytest.raises(DuplicateEmailError):
        await create_user(
            "[email protected]",
            name="Bob",
            national_id="87654321",
            users=users,
            audit=audit,
            secrets=secrets,
            now=_now(),
        )


async def test_create_user_rejects_empty_name() -> None:
    users = FakeUserRepository()
    audit = FakeAuditLog()
    secrets = FakeSecretManager()
    with pytest.raises(ValueError, match="non-empty"):
        await create_user(
            "[email protected]",
            name="   ",
            national_id="12345678",
            users=users,
            audit=audit,
            secrets=secrets,
            now=_now(),
        )


# ---------------------------------------------------------------------------
# disable_user
# ---------------------------------------------------------------------------


async def test_disable_user_transitions_active_to_disabled() -> None:
    users = FakeUserRepository()
    users.add(_user())
    audit = FakeAuditLog()
    await disable_user(
        list(users.by_id.keys())[0],
        users=users,
        audit=audit,
        now=_now(),
    )
    assert users.by_id[list(users.by_id.keys())[0]].status is UserStatus.DISABLED
    assert len(audit.entries) == 1
    assert audit.entries[0].event_type == "users.disable"


async def test_disable_user_rejects_unknown_id() -> None:
    users = FakeUserRepository()
    audit = FakeAuditLog()
    with pytest.raises(UserNotFoundError):
        await disable_user(
            __import__("uuid").uuid4(),
            users=users,
            audit=audit,
            now=_now(),
        )


async def test_disable_user_rejects_non_active_status() -> None:
    users = FakeUserRepository()
    user = _user(status=UserStatus.DISABLED)
    users.add(user)
    audit = FakeAuditLog()
    with pytest.raises(ValueError, match="status=active"):
        await disable_user(user.id, users=users, audit=audit, now=_now())


# ---------------------------------------------------------------------------
# grant_global_admin
# ---------------------------------------------------------------------------


async def test_grant_global_admin_creates_row_and_audits() -> None:
    users = FakeUserRepository()
    users.add(_user())
    audit = FakeAuditLog()
    global_admins = FakeGlobalAdminRepository()
    user_id = list(users.by_id.keys())[0]
    await grant_global_admin(
        user_id,
        users=users,
        global_admins=global_admins,
        audit=audit,
        now=_now(),
    )
    assert await global_admins.is_global_admin(user_id)
    assert len(audit.entries) == 1


async def test_grant_global_admin_is_idempotent() -> None:
    users = FakeUserRepository()
    users.add(_user())
    audit = FakeAuditLog()
    global_admins = FakeGlobalAdminRepository()
    user_id = list(users.by_id.keys())[0]
    await grant_global_admin(
        user_id, users=users, global_admins=global_admins, audit=audit, now=_now()
    )
    await grant_global_admin(
        user_id, users=users, global_admins=global_admins, audit=audit, now=_now()
    )
    # Second call is a no-op: no extra audit row.
    assert len(audit.entries) == 1


# ---------------------------------------------------------------------------
# revoke_global_admin
# ---------------------------------------------------------------------------


async def test_revoke_global_admin_removes_row_and_audits() -> None:
    users = FakeUserRepository()
    users.add(_user("[email protected]"))
    users.add(_user("[email protected]"))
    audit = FakeAuditLog()
    global_admins = FakeGlobalAdminRepository()
    alice = list(users.by_id.keys())[0]
    bob = list(users.by_id.keys())[1]
    await global_admins.grant(alice)
    await global_admins.grant(bob)
    await revoke_global_admin(
        bob,
        users=users,
        global_admins=global_admins,
        audit=audit,
        now=_now(),
    )
    assert not await global_admins.is_global_admin(bob)
    assert await global_admins.is_global_admin(alice)
    assert len(audit.entries) == 1
    assert audit.entries[0].event_type == "global_admins.revoke"


async def test_revoke_global_admin_rejects_last_admin() -> None:
    users = FakeUserRepository()
    users.add(_user())
    audit = FakeAuditLog()
    global_admins = FakeGlobalAdminRepository()
    user_id = list(users.by_id.keys())[0]
    await global_admins.grant(user_id)
    with pytest.raises(ValueError, match="zero admins"):
        await revoke_global_admin(
            user_id, users=users, global_admins=global_admins, audit=audit, now=_now()
        )
    assert audit.entries[-1].result == "rejected"


# ---------------------------------------------------------------------------
# assign_profile
# ---------------------------------------------------------------------------


async def test_assign_profile_creates_assignment() -> None:
    from uuid import uuid4

    users = FakeUserRepository()
    users.add(_user())
    apps = FakeAppRepository()
    profiles = FakeProfileRepository()
    assignments = FakeAssignmentRepository()
    audit = FakeAuditLog()
    user_id = list(users.by_id.keys())[0]
    app = App(
        id=1,
        name="Lanzadera",
        short_code="lanza",
        deployment_topology="central",
        requires_office_presence=False,
        registration_status="active",
        created_at=_now(),
        updated_at=_now(),
    )
    apps.add(app)
    profile = Profile(
        id=uuid4(),
        app_id=1,
        code="admin",
        name="Admin",
        capabilities={},
        active=True,
        created_at=_now(),
        updated_at=_now(),
    )
    profiles.add(profile)
    await assign_profile(
        user_id,
        1,
        profile.id,
        users=users,
        apps=apps,
        profiles=profiles,
        assignments=assignments,
        audit=audit,
        now=_now(),
    )
    assert user_id in [a.user_id for a in assignments.by_user.get(user_id, [])]


async def test_assign_profile_rejects_sinacceso_exclusivity() -> None:
    from uuid import uuid4

    users = FakeUserRepository()
    users.add(_user())
    apps = FakeAppRepository()
    profiles = FakeProfileRepository()
    assignments = FakeAssignmentRepository()
    audit = FakeAuditLog()
    user_id = list(users.by_id.keys())[0]
    apps.add(
        App(
            id=1,
            name="A",
            short_code="a",
            deployment_topology="central",
            requires_office_presence=False,
            registration_status="active",
            created_at=_now(),
            updated_at=_now(),
        )
    )
    # Existing non-SinAcceso assignment
    other_profile = Profile(
        id=uuid4(),
        app_id=1,
        code="power",
        name="Power",
        capabilities={"all": True},
        active=True,
        created_at=_now(),
        updated_at=_now(),
    )
    profiles.add(other_profile)
    await assignments.create(user_id, 1, other_profile.id)
    # Try to assign SinAcceso
    sinacceso = Profile(
        id=uuid4(),
        app_id=1,
        code="no",
        name="No",
        capabilities={"SinAcceso": True},
        active=True,
        created_at=_now(),
        updated_at=_now(),
    )
    profiles.add(sinacceso)
    with pytest.raises(ValueError, match="cannot assign SinAcceso"):
        await assign_profile(
            user_id,
            1,
            sinacceso.id,
            users=users,
            apps=apps,
            profiles=profiles,
            assignments=assignments,
            audit=audit,
            now=_now(),
        )


# ---------------------------------------------------------------------------
# list_effective_apps
# ---------------------------------------------------------------------------


async def test_list_effective_apps_returns_live_only() -> None:
    from uuid import uuid4

    users = FakeUserRepository()
    users.add(_user())
    apps = FakeAppRepository()
    assignments = FakeAssignmentRepository()
    active_app = App(
        id=1,
        name="Active",
        short_code="a",
        deployment_topology="central",
        requires_office_presence=False,
        registration_status="active",
        created_at=_now(),
        updated_at=_now(),
    )
    pending_app = App(
        id=2,
        name="Pending",
        short_code="p",
        deployment_topology="central",
        requires_office_presence=False,
        registration_status="pending",
        created_at=_now(),
        updated_at=_now(),
    )
    apps.add(active_app)
    apps.add(pending_app)
    user_id = list(users.by_id.keys())[0]
    await assignments.create(user_id, 1, uuid4())
    result = await list_effective_apps(user_id, apps=apps, assignments=assignments)
    # active_app is live AND assigned; pending_app is not live.
    assert active_app in result
    assert pending_app not in result


# ---------------------------------------------------------------------------
# audit_append
# ---------------------------------------------------------------------------


async def test_audit_append_stamps_correlation_id() -> None:
    audit = FakeAuditLog()
    event = await audit_append(
        "test.event",
        actor_id=None,
        target_id="t1",
        result="success",
        payload={"k": "v"},
        audit=audit,
        now=_now(),
    )
    assert event.event_type == "test.event"
    assert event.correlation_id is not None
    assert event.module == "lanzadera"


# ---------------------------------------------------------------------------
# bootstrap_global_admins
# ---------------------------------------------------------------------------


async def test_bootstrap_global_admins_idempotent(monkeypatch) -> None:
    monkeypatch.setenv("GLOBAL_ADMIN_EMAILS", "a-admin@example.com;b-admin@example.com")
    users = FakeUserRepository()
    global_admins = FakeGlobalAdminRepository()
    audit = FakeAuditLog()
    secrets = FakeSecretManager()
    created_first = await bootstrap_global_admins(
        users=users,
        global_admins=global_admins,
        audit=audit,
        secrets=secrets,
        now=_now(),
    )
    assert created_first == 2
    assert len(users.by_id) == 2
    created_second = await bootstrap_global_admins(
        users=users,
        global_admins=global_admins,
        audit=audit,
        secrets=secrets,
        now=_now(),
    )
    assert created_second == 0  # second run is no-op


async def test_bootstrap_global_admins_no_env_var() -> None:
    monkeypatch = pytest.MonkeyPatch()
    monkeypatch.delenv("GLOBAL_ADMIN_EMAILS", raising=False)
    users = FakeUserRepository()
    global_admins = FakeGlobalAdminRepository()
    audit = FakeAuditLog()
    secrets = FakeSecretManager()
    created = await bootstrap_global_admins(
        users=users,
        global_admins=global_admins,
        audit=audit,
        secrets=secrets,
        now=_now(),
    )
    assert created == 0
    assert len(users.by_id) == 0


# ---------------------------------------------------------------------------
# set_password
# ---------------------------------------------------------------------------


async def test_set_password_creates_and_activates_first_admin() -> None:
    users = FakeUserRepository()
    global_admins = FakeGlobalAdminRepository()
    audit = FakeAuditLog()
    secrets = FakeSecretManager()
    hasher = FakeCredentialHasher()
    await set_password(
        "[email protected]",
        new_password="supersecret",
        hasher=hasher,
        users=users,
        global_admins=global_admins,
        audit=audit,
        secrets=secrets,
        now=_now(),
    )
    assert "[email protected]" in users.by_email
    user = users.by_email["[email protected]"]
    assert user.status is UserStatus.ACTIVE
    assert user.password_hash == "fake:supersecret"
    assert await global_admins.is_global_admin(user.id)


async def test_set_password_existing_user_idempotent() -> None:
    users = FakeUserRepository()
    users.add(_user())
    global_admins = FakeGlobalAdminRepository()
    audit = FakeAuditLog()
    secrets = FakeSecretManager()
    hasher = FakeCredentialHasher()
    user_id = list(users.by_id.keys())[0]
    await global_admins.grant(user_id)
    await set_password(
        "[email protected]",
        new_password="newpw",
        hasher=hasher,
        users=users,
        global_admins=global_admins,
        audit=audit,
        secrets=secrets,
        now=_now(),
    )
    user = users.by_id[user_id]
    assert user.password_hash == "fake:newpw"
