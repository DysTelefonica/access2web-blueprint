# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W-TEST (#519)
"""Container-level integration tests for ``LanzaderaContainer`` (issue #519).

Wires the in-memory fakes into a real ``LanzaderaContainer`` and asserts
the use-case partials resolve end-to-end: create / disable / assign
each flow through ``container.<use_case>`` → ``use_case(...)`` →
fake repository → fake audit log.

The tests cover the W54-W59 slice:
- ``list_all_users`` paginates rows + total (W59, #517).
- ``create_user`` mints a UUID, encrypts the DNI, persists the row,
  appends the audit row (W54, #43).
- ``disable_user`` flips ACTIVE → DISABLED and emits the audit row.
- ``assign_profile`` creates the (user, app, profile) row and emits
  the audit row.
"""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import UUID, uuid4

import pytest

from app.src.modules.lanzadera.domain.app import App, AppRegistrationStatus, AppTopology
from app.src.modules.lanzadera.domain.errors import DuplicateEmailError, UserNotFoundError
from app.src.modules.lanzadera.domain.profile import Profile
from app.src.modules.lanzadera.domain.user import User, UserStatus
from tests.lanzadera.conftest import FakeFixtures

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _seed_active_user(fakes: FakeFixtures, *, email: str = "alice@enterprise.test") -> User:
    """Seed a user with ``status=ACTIVE`` so ``disable_user`` accepts the transition."""
    user = User(
        id=uuid4(),
        email=email,
        name="Alice",
        dni_encrypted=b"enc:11111111",
        password_hash="fake:pass",
        status=UserStatus.ACTIVE,
        failed_attempts=0,
        last_login_at=None,
        created_at=datetime.now(UTC),
        updated_at=datetime.now(UTC),
    )
    fakes.users.add(user)
    return user


def _seed_app(fakes: FakeFixtures, *, app_id: int = 1) -> App:
    """Seed an active app the assign_profile flow can target."""
    app = App(
        id=app_id,
        name="Expedientes",
        short_code="EXP",
        deployment_topology=AppTopology.CENTRAL,
        requires_office_presence=False,
        registration_status=AppRegistrationStatus.ACTIVE,
        created_at=datetime.now(UTC),
        updated_at=datetime.now(UTC),
    )
    fakes.apps.add(app)
    return app


def _seed_profile(fakes: FakeFixtures, *, app_id: int = 1) -> Profile:
    """Seed a profile that targets ``app_id`` with a non-SinAcceso capability set."""
    profile = Profile(
        id=uuid4(),
        app_id=app_id,
        code="default",
        name="Default",
        capabilities={"Read": True, "Write": True},
        active=True,
        created_at=datetime.now(UTC),
        updated_at=datetime.now(UTC),
    )
    fakes.profiles.add(profile)
    return profile


# ---------------------------------------------------------------------------
# list_all_users — W59 (#517) pagination
# ---------------------------------------------------------------------------


async def test_list_all_users_returns_tuple(container, fake_fixtures: FakeFixtures):
    """``list_all_users`` returns ``(rows, total)`` per the W59 contract."""
    _seed_active_user(fake_fixtures, email="alice@enterprise.test")
    _seed_active_user(fake_fixtures, email="bob@enterprise.test")
    _seed_active_user(fake_fixtures, email="carol@enterprise.test")

    rows, total = await container.list_all_users(limit=50, offset=0)

    assert total == 3
    assert len(rows) == 3


async def test_list_all_users_respects_limit_and_offset(container, fake_fixtures: FakeFixtures):
    """Pagination slices the rows; ``total`` reports the full count."""
    _seed_active_user(fake_fixtures, email="alice@enterprise.test")
    _seed_active_user(fake_fixtures, email="bob@enterprise.test")
    _seed_active_user(fake_fixtures, email="carol@enterprise.test")

    rows_page_1, total = await container.list_all_users(limit=2, offset=0)
    rows_page_2, total_again = await container.list_all_users(limit=2, offset=2)

    assert total == 3
    assert total_again == 3
    assert len(rows_page_1) == 2
    assert len(rows_page_2) == 1


# ---------------------------------------------------------------------------
# create_user — W54, D5, D7, D89, DA-11
# ---------------------------------------------------------------------------


async def test_create_user_returns_user_with_uuid(container, fake_fixtures: FakeFixtures):
    """``create_user`` mints a UUID and persists the row with the normalised email."""
    new_user = await container.create_user(
        email="  alice" + "@" + "enterprise.test  ",
        name="Alice",
        national_id="11111111",
    )

    assert isinstance(new_user.id, UUID)
    assert new_user.email == "alice" + "@" + "enterprise.test"
    assert new_user.name == "Alice"
    assert new_user.status is UserStatus.PASSWORD_RESET_REQUIRED
    assert new_user.password_hash is None
    # The fake persists the row the use case created.
    alice_email = "alice" + "@" + "enterprise.test"
    assert fake_fixtures.users.by_email[alice_email] is new_user


async def test_create_user_encrypts_dni(container, fake_fixtures: FakeFixtures):
    """The DNI lands as ciphertext via the ``SecretManager.encrypt`` seam."""
    await container.create_user(
        email="alice" + "@" + "enterprise.test",
        name="Alice",
        national_id="11111111",
    )

    assert fake_fixtures.secrets.encrypted == ["11111111"]
    persisted = fake_fixtures.users.by_email["alice" + "@" + "enterprise.test"]
    assert persisted.dni_encrypted == b"enc:11111111"


async def test_create_user_audit_row_created(container, fake_fixtures: FakeFixtures):
    """``create_user`` appends exactly one ``users.create`` audit row (DA-11)."""
    actor_id = uuid4()
    await container.create_user(
        email="alice" + "@" + "enterprise.test",
        name="Alice",
        national_id="11111111",
        actor_id=actor_id,
    )

    assert len(fake_fixtures.audit.entries) == 1
    entry = fake_fixtures.audit.entries[0]
    assert entry.event_type == "users.create"
    assert entry.actor_id == actor_id
    assert entry.result == "success"
    assert entry.payload["email"] == "alice" + "@" + "enterprise.test"


async def test_create_user_raises_on_duplicate_email(container, fake_fixtures: FakeFixtures):
    """A second ``create_user`` with the same email raises ``DuplicateEmailError`` (D7)."""
    _seed_active_user(fake_fixtures, email="alice" + "@" + "enterprise.test")

    with pytest.raises(DuplicateEmailError):
        await container.create_user(
            email="alice" + "@" + "enterprise.test",
            name="Bob",
            national_id="22222222",
        )


async def test_create_user_does_not_audit_on_duplicate(container, fake_fixtures: FakeFixtures):
    """No audit row on the duplicate path — DA-11 forbids side effects on rejection."""
    _seed_active_user(fake_fixtures, email="alice" + "@" + "enterprise.test")

    with pytest.raises(DuplicateEmailError):
        await container.create_user(
            email="alice" + "@" + "enterprise.test",
            name="Bob",
            national_id="22222222",
        )

    assert fake_fixtures.audit.entries == []


# ---------------------------------------------------------------------------
# disable_user — D42, D89, DA-11
# ---------------------------------------------------------------------------


async def test_disable_user_sets_status_disabled(container, fake_fixtures: FakeFixtures):
    """``disable_user`` flips ACTIVE → DISABLED on the persisted row."""
    user = _seed_active_user(fake_fixtures)

    await container.disable_user(user_id=user.id)

    assert user.status is UserStatus.DISABLED
    assert fake_fixtures.users.status_calls[-1] == (user.id, UserStatus.DISABLED)


async def test_disable_user_audit_logged(container, fake_fixtures: FakeFixtures):
    """``disable_user`` appends exactly one ``users.disable`` audit row."""
    user = _seed_active_user(fake_fixtures)

    await container.disable_user(user_id=user.id)

    assert len(fake_fixtures.audit.entries) == 1
    entry = fake_fixtures.audit.entries[0]
    assert entry.event_type == "users.disable"
    assert entry.payload == {"from": "active", "to": "disabled"}


async def test_disable_user_raises_for_unknown_user(container, fake_fixtures: FakeFixtures):
    """``disable_user`` raises ``UserNotFoundError`` for a non-existent id."""
    with pytest.raises(UserNotFoundError):
        await container.disable_user(user_id=uuid4())


async def test_disable_user_rejects_non_active_user(container, fake_fixtures: FakeFixtures):
    """``disable_user`` rejects users whose current status is not ACTIVE (D42)."""
    user = User(
        id=uuid4(),
        email="disabled@enterprise.test",
        name="Dave",
        dni_encrypted=b"enc:33333333",
        password_hash="fake:pass",
        status=UserStatus.DISABLED,
        failed_attempts=0,
        last_login_at=None,
        created_at=datetime.now(UTC),
        updated_at=datetime.now(UTC),
    )
    fake_fixtures.users.add(user)

    with pytest.raises(ValueError, match="disable_user requires status=active"):
        await container.disable_user(user_id=user.id)


# ---------------------------------------------------------------------------
# assign_profile — D22, DA-12, H11, DA-11
# ---------------------------------------------------------------------------


async def test_assign_profile_creates_assignment(container, fake_fixtures: FakeFixtures):
    """``assign_profile`` persists the (user, app, profile) triple on the fake repo."""
    user = _seed_active_user(fake_fixtures)
    _seed_app(fake_fixtures, app_id=42)
    profile = _seed_profile(fake_fixtures, app_id=42)

    await container.assign_profile(
        user_id=user.id,
        app_id=42,
        profile_id=profile.id,
    )

    assert len(fake_fixtures.assignments.create_calls) == 1
    called = fake_fixtures.assignments.create_calls[0]
    assert called[0] == user.id
    assert called[1] == 42


async def test_assign_profile_audit_logged(container, fake_fixtures: FakeFixtures):
    """``assign_profile`` appends exactly one ``assignments.create`` audit row."""
    user = _seed_active_user(fake_fixtures)
    _seed_app(fake_fixtures, app_id=42)
    profile = _seed_profile(fake_fixtures, app_id=42)

    await container.assign_profile(
        user_id=user.id,
        app_id=42,
        profile_id=profile.id,
    )

    assert len(fake_fixtures.audit.entries) == 1
    entry = fake_fixtures.audit.entries[0]
    assert entry.event_type == "assignments.create"
    assert entry.payload["app_id"] == 42
    assert entry.payload["profile_id"] == str(profile.id)


async def test_assign_profile_rejects_unknown_user(container, fake_fixtures: FakeFixtures):
    """``assign_profile`` raises ``UserNotFoundError`` for a non-existent user."""
    _seed_app(fake_fixtures, app_id=42)
    profile = _seed_profile(fake_fixtures, app_id=42)

    with pytest.raises(UserNotFoundError):
        await container.assign_profile(
            user_id=uuid4(),
            app_id=42,
            profile_id=profile.id,
        )


# ---------------------------------------------------------------------------
# bootstrap_global_admins — D91, idempotent env-driven seed
# ---------------------------------------------------------------------------


async def test_bootstrap_global_admins_returns_count(
    container,
    fake_fixtures: FakeFixtures,
    monkeypatch: pytest.MonkeyPatch,
):
    """``bootstrap_global_admins`` returns the number of new admins granted.

    Two emails are seeded via ``GLOBAL_ADMIN_EMAILS``: neither user
    exists, so the use case creates both rows and grants them global
    admin in a single pass — the return value is ``2`` (D91: count of
    NEW rows, not the running total).
    """
    # Built via concatenation so the test source never embeds an ``@``
    # pattern that a doc-renderer would auto-obfuscate.
    alpha = "alpha" + "@" + "enterprise.test"
    beta = "beta" + "@" + "enterprise.test"
    monkeypatch.setenv("GLOBAL_ADMIN_EMAILS", f"{alpha};{beta}")

    created = await container.bootstrap_global_admins()

    assert created == 2
    # The users were minted through the fake repository with the
    # placeholder-DNI ciphertext and the password-reset-required status.
    assert alpha in fake_fixtures.users.by_email
    assert beta in fake_fixtures.users.by_email
    for email in (alpha, beta):
        user = fake_fixtures.users.by_email[email]
        assert user.dni_encrypted == b"enc:BOOTSTRAP-PLACEHOLDER"
        assert user.status is UserStatus.PASSWORD_RESET_REQUIRED
    # Both emails were granted — the fake's grant_calls list is the
    # exact count.
    assert len(fake_fixtures.global_admins.grant_calls) == 2
    # DA-11: one audit row per new admin.
    assert len(fake_fixtures.audit.entries) == 2
    for entry in fake_fixtures.audit.entries:
        assert entry.event_type == "global_admins.bootstrap"


async def test_bootstrap_global_admins_is_idempotent(
    container,
    fake_fixtures: FakeFixtures,
    monkeypatch: pytest.MonkeyPatch,
):
    """Re-running with the same emails returns ``0`` (D91 idempotency).

    After the first call both users are already global admins, so the
    second pass issues zero ``grant`` calls and emits zero audit rows.
    """
    alpha = "alpha" + "@" + "enterprise.test"
    monkeypatch.setenv("GLOBAL_ADMIN_EMAILS", alpha)

    first = await container.bootstrap_global_admins()
    second = await container.bootstrap_global_admins()

    assert first == 1
    assert second == 0
    # The fake's grant_calls list still carries the one call from the
    # first pass — the second pass did not re-issue it.
    assert len(fake_fixtures.global_admins.grant_calls) == 1
    # The audit log has exactly one entry — the second pass skipped
    # the append because the membership was already in place.
    assert len(fake_fixtures.audit.entries) == 1


async def test_bootstrap_global_admins_noop_when_env_unset(
    container,
    fake_fixtures: FakeFixtures,
    monkeypatch: pytest.MonkeyPatch,
):
    """Unset ``GLOBAL_ADMIN_EMAILS`` is a no-op: returns ``0``, no rows touched."""
    monkeypatch.delenv("GLOBAL_ADMIN_EMAILS", raising=False)

    created = await container.bootstrap_global_admins()

    assert created == 0
    assert fake_fixtures.global_admins.grant_calls == []
    assert fake_fixtures.audit.entries == []
