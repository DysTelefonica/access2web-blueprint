"""RED-first tests for the platform-user CLI (WU DL3, issue #56).

The tests cover the contract that the subcommands return the right
``CommandResult`` (success / failure / cancelled) and that the
destructive ones (set-password, grant-global-admin, revoke-global-
admin, assign-profile) honour the D26 confirmation prompt via the
``confirmed=True`` flag (the production driver passes ``False`` and
the CLI helper prompts; the tests pass ``True`` to skip the prompt).

The tests do not touch the live database; they use in-memory fakes
that satisfy the relevant ``Protocol`` shapes (UserRepository,
GlobalAdminRepository, AssignmentRepository, AppRepository,
ProfileRepository). The hasher is the real ``CredentialHasherArgon2id``
because it is pure-stdlib and fast.
"""

from __future__ import annotations

import dataclasses
from uuid import UUID, uuid4

import pytest

from app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id import (
    CredentialHasherArgon2id,
)
from app.src.modules.lanzadera.delivery.cli.platform_user import (
    build_parser,
    cmd_assign_profile,
    cmd_grant_global_admin,
    cmd_list_apps,
    cmd_revoke_global_admin,
    cmd_set_password,
)
from app.src.modules.lanzadera.domain.ports import UserRepository
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)
from app.src.modules.lanzadera.domain.user import User, UserStatus

# ---------------------------------------------------------------------------
# In-memory fakes. Each fake satisfies the structural Protocol the
# corresponding cmd_* function needs. The fakes record every call so
# the tests can assert that the destructive commands issued exactly
# one write per "yes" answer.
# ---------------------------------------------------------------------------


class _FakeUserRepo(UserRepository):
    def __init__(self, user: User | None = None) -> None:  # type: ignore[no-untyped-def]
        self._user = user
        self.update_calls: list[tuple[UUID, str]] = []

    async def get_by_email(self, email: str) -> User | None:
        if self._user is None:
            return None
        if self._user.email == email:
            return self._user
        return None

    async def get_by_id(self, user_id: UUID) -> User | None:  # noqa: ARG002
        if self._user is not None and self._user.id == user_id:
            return self._user
        return None

    async def update_password_and_activate(self, user_id: UUID, password_hash: str) -> None:
        self.update_calls.append((user_id, password_hash))
        if self._user is not None:
            self._user = dataclasses.replace(
                self._user,
                password_hash=password_hash,
                status=UserStatus.ACTIVE,
            )

    async def update_status(self, user_id: UUID, status: UserStatus) -> None:  # noqa: ARG002
        return None

    async def list_all(self) -> list[User]:  # noqa: ARG002
        return [] if self._user is None else [self._user]

    async def create(self, user: User) -> None:  # noqa: ARG002
        self._user = user


class _FakeGlobalAdminRepo(GlobalAdminRepositoryPort):
    def __init__(self, members: set[UUID] | None = None) -> None:
        self._members: set[UUID] = set(members) if members is not None else set()
        self.grant_calls: list[UUID] = []
        self.revoke_calls: list[UUID] = []

    async def list_all(self) -> list[object]:  # noqa: ARG002
        return list(self._members)

    async def is_global_admin(self, user_id: UUID) -> bool:
        return user_id in self._members

    async def grant(self, user_id: UUID) -> None:
        if user_id in self._members:
            raise ValueError(f"user {user_id} is already a global admin")
        self.grant_calls.append(user_id)
        self._members.add(user_id)

    async def revoke(self, user_id: UUID) -> None:
        if user_id not in self._members:
            return
        if len(self._members) == 1:
            raise ValueError("cannot revoke the last global admin")
        self.revoke_calls.append(user_id)
        self._members.discard(user_id)


@dataclasses.dataclass
class _App:
    id: int
    name: str


class _FakeAppRepo:
    def __init__(self, apps: list[_App] | None = None) -> None:
        self._apps = apps or [_App(id=1, name="Expedientes")]

    async def list_active(self) -> list[_App]:
        return list(self._apps)


@dataclasses.dataclass
class _Profile:
    id: UUID
    code: str


class _FakeProfileRepo:
    def __init__(self, profiles: list[_Profile] | None = None) -> None:
        self._profiles = profiles or [_Profile(id=uuid4(), code="default")]

    async def get_by_code(self, app_id: int, code: str) -> _Profile | None:
        for p in self._profiles:
            if p.code == code:
                return p
        return None


class _FakeAssignmentRepo(AssignmentRepositoryPort):
    def __init__(self) -> None:
        self.created: list[tuple[UUID, int, UUID]] = []

    async def create(self, user_id: UUID, app_id: int, profile_id: UUID) -> None:
        self.created.append((user_id, app_id, profile_id))


# ---------------------------------------------------------------------------
# Fixtures.
# ---------------------------------------------------------------------------


@pytest.fixture
def hasher() -> CredentialHasherArgon2id:
    return CredentialHasherArgon2id()


@pytest.fixture
def sample_user() -> User:
    return User(
        id=uuid4(),
        email="[email protected]",
        name="Alice",
        dni_encrypted=None,
        password_hash=None,
        status=UserStatus.PASSWORD_RESET_REQUIRED,
        failed_attempts=0,
        last_login_at=None,
        created_at=__import__("datetime").datetime.now(__import__("datetime").UTC),
        updated_at=__import__("datetime").datetime.now(__import__("datetime").UTC),
    )


# ---------------------------------------------------------------------------
# set-password.
# ---------------------------------------------------------------------------


async def test_set_password_marks_user_active_and_hashes(
    hasher: CredentialHasherArgon2id, sample_user: User, monkeypatch: pytest.MonkeyPatch
) -> None:
    # Override the password reader to bypass the tty dependency in the
    # pytest capture environment. The CLI uses ``_read_password`` from
    # the platform_user module.
    from app.src.modules.lanzadera.delivery.cli import platform_user

    monkeypatch.setattr(platform_user, "_read_password", lambda prompt: "new-password")
    users = _FakeUserRepo(sample_user)
    result = await cmd_set_password("[email protected]", users, hasher, confirmed=True)
    assert result.exit_code == 0
    assert len(users.update_calls) == 1
    user_id, recorded_hash = users.update_calls[0]
    assert user_id == sample_user.id
    assert recorded_hash != "new-password"  # the cmd hashed the password
    assert recorded_hash != ""  # and wrote a non-empty hash
    assert recorded_hash.startswith("$argon2id$")


async def test_set_password_returns_error_for_unknown_user(
    hasher: CredentialHasherArgon2id,
) -> None:
    users = _FakeUserRepo()  # no user at all
    result = await cmd_set_password("[email protected]", users, hasher, confirmed=True)
    assert result.exit_code == 1
    assert "user not found" in result.message


# ---------------------------------------------------------------------------
# grant-global-admin.
# ---------------------------------------------------------------------------


async def test_grant_global_admin_inserts_row(
    sample_user: User,
) -> None:
    users = _FakeUserRepo(sample_user)
    admins = _FakeGlobalAdminRepo()
    result = await cmd_grant_global_admin("[email protected]", users, admins, confirmed=True)
    assert result.exit_code == 0
    assert admins.grant_calls == [sample_user.id]
    assert await admins.is_global_admin(sample_user.id) is True


async def test_grant_global_admin_fails_on_already_admin(
    sample_user: User,
) -> None:
    users = _FakeUserRepo(sample_user)
    admins = _FakeGlobalAdminRepo(members={sample_user.id})
    result = await cmd_grant_global_admin("[email protected]", users, admins, confirmed=True)
    assert result.exit_code == 1
    assert "cannot grant" in result.message


# ---------------------------------------------------------------------------
# revoke-global-admin.
# ---------------------------------------------------------------------------


async def test_revoke_global_admin_removes_row(
    sample_user: User,
) -> None:
    other = uuid4()
    users = _FakeUserRepo(sample_user)
    admins = _FakeGlobalAdminRepo(members={sample_user.id, other})
    result = await cmd_revoke_global_admin("[email protected]", users, admins, confirmed=True)
    assert result.exit_code == 0
    assert admins.revoke_calls == [sample_user.id]
    assert await admins.is_global_admin(sample_user.id) is False
    assert await admins.is_global_admin(other) is True


async def test_revoke_global_admin_fails_on_last_admin(
    sample_user: User,
) -> None:
    users = _FakeUserRepo(sample_user)
    admins = _FakeGlobalAdminRepo(members={sample_user.id})
    result = await cmd_revoke_global_admin("[email protected]", users, admins, confirmed=True)
    assert result.exit_code == 1
    assert "cannot revoke" in result.message
    assert "last global admin" in result.message


# ---------------------------------------------------------------------------
# list-apps.
# ---------------------------------------------------------------------------


async def test_list_apps_prints_active() -> None:
    apps = _FakeAppRepo([_App(id=1, name="Expedientes"), _App(id=2, name="HPS")])
    result = await cmd_list_apps(apps)
    assert result.exit_code == 0
    assert "Expedientes" in result.message
    assert "HPS" in result.message
    assert "active apps" in result.message


# ---------------------------------------------------------------------------
# assign-profile.
# ---------------------------------------------------------------------------


async def test_assign_profile_creates_row(
    sample_user: User,
) -> None:
    users = _FakeUserRepo(sample_user)
    profile = _Profile(id=uuid4(), code="default")
    profiles = _FakeProfileRepo([profile])
    assignments = _FakeAssignmentRepo()

    result = await cmd_assign_profile(
        "[email protected]",
        app_id=42,
        profile_code="default",
        users=users,
        profiles=profiles,
        assignments=assignments,
        confirmed=True,
    )
    assert result.exit_code == 0
    assert assignments.created == [(sample_user.id, 42, profile.id)]


async def test_assign_profile_fails_for_unknown_profile(
    sample_user: User,
) -> None:
    users = _FakeUserRepo(sample_user)
    profiles = _FakeProfileRepo([])  # no profile in the app
    assignments = _FakeAssignmentRepo()

    result = await cmd_assign_profile(
        "[email protected]",
        app_id=42,
        profile_code="nope",
        users=users,
        profiles=profiles,
        assignments=assignments,
        confirmed=True,
    )
    assert result.exit_code == 1
    assert "not found" in result.message


# ---------------------------------------------------------------------------
# CLI parser surface.
# ---------------------------------------------------------------------------


def test_parser_has_all_five_subcommands() -> None:
    parser = build_parser()
    parsed = parser.parse_args(["set-password", "[email protected]"])
    assert parsed.subcommand == "set-password"
    assert parsed.email == "[email protected]"

    parsed = parser.parse_args(["grant-global-admin", "[email protected]"])
    assert parsed.subcommand == "grant-global-admin"

    parsed = parser.parse_args(["revoke-global-admin", "[email protected]"])
    assert parsed.subcommand == "revoke-global-admin"

    parsed = parser.parse_args(["list-apps"])
    assert parsed.subcommand == "list-apps"

    parsed = parser.parse_args(["assign-profile", "[email protected]", "42", "default"])
    assert parsed.subcommand == "assign-profile"
    assert parsed.app_id == 42
    assert parsed.profile_code == "default"


def test_parser_requires_subcommand() -> None:
    """DA-5: the destructive subcommands MUST be invoked explicitly; the
    parser requires ``subcommand`` so a bare ``gentle-ai platform user``
    fails with a usage error (no implicit default to set-password)."""
    parser = build_parser()
    with pytest.raises(SystemExit):
        parser.parse_args([])
