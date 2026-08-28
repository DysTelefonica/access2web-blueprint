"""Contract tests for the admin HTTP delivery (DL2, issue #55).

The tests use ``httpx.AsyncClient`` + ``ASGITransport`` so the
router runs in-process without a live server. The use-case ports are
fakes that satisfy the relevant ``Protocol`` shapes; the router
treats them as the production adapters.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from fastapi import FastAPI
from fastapi.templating import Jinja2Templates
from httpx import ASGITransport, AsyncClient

from app.src.modules.lanzadera.delivery.http.admin import build_router
from app.src.modules.lanzadera.domain.app import App
from app.src.modules.lanzadera.domain.user import User, UserStatus

# ---------------------------------------------------------------------------
# Fakes — one per port the router depends on. Each fake records every
# call so the tests can assert that destructive operations issued
# exactly one write per call.
# ---------------------------------------------------------------------------


class FakeUserRepository:
    def __init__(self) -> None:
        self._users: dict[str, User] = {}
        self.updates: list[tuple[object, UserStatus]] = []

    async def get_by_email(self, email: str) -> User | None:
        return self._users.get(email)

    async def update_status(self, user_id: object, status: UserStatus) -> None:
        self.updates.append((user_id, status))
        for u in self._users.values():
            if u.id == user_id:
                u.status = status
                return

    async def create(self, user: User) -> None:
        self._users[user.email] = user

    async def list_all(self) -> list[User]:
        return list(self._users.values())


class FakeAppRepository:
    def __init__(self) -> None:
        self._apps: list[App] = []

    async def list_active(self) -> list[App]:
        return list(self._apps)


class FakeAssignmentRepository:
    def __init__(self) -> None:
        self.created: list[tuple[object, int, object]] = []

    async def create(self, user_id: object, app_id: int, profile_id: object) -> None:
        self.created.append((user_id, app_id, profile_id))


class FakeAuditLogPort:
    def __init__(self) -> None:
        self._events: list[object] = []

    async def list_for_actor(self, actor_id: object, since: object) -> list[object]:
        return list(self._events)


class FakeGlobalAdminRepository:
    async def is_global_admin(self, user_id: object) -> bool:
        return True


def fake_now():
    from datetime import UTC, datetime

    return datetime(2026, 8, 21, 21, 0, 0, tzinfo=UTC)


def fake_uuid():
    from uuid import uuid4

    return uuid4()


# ---------------------------------------------------------------------------
# Fixture: build a FastAPI app with the admin router wired to the fakes.
# ---------------------------------------------------------------------------


@pytest.fixture
def fakes():
    return {
        "users": FakeUserRepository(),
        "apps": FakeAppRepository(),
        "assignments": FakeAssignmentRepository(),
        "audit": FakeAuditLogPort(),
        "admins": FakeGlobalAdminRepository(),
    }


@pytest.fixture
def app(fakes):
    from app.src.modules.lanzadera.delivery.http import templates

    templates_dir = Path(templates.__file__).parent
    templates_jinja = Jinja2Templates(directory=str(templates_dir))
    application = FastAPI()
    application.include_router(
        build_router(
            templates=templates_jinja,
            users=fakes["users"],
            apps=fakes["apps"],
            assignments=fakes["assignments"],
            audit=fakes["audit"],
            admins=fakes["admins"],
        )
    )
    return application, fakes


# ---------------------------------------------------------------------------
# Tests.
# ---------------------------------------------------------------------------


async def test_list_users_renders_empty_table(client: AsyncClient) -> None:
    r = await client.get("/admin/users")
    assert r.status_code == 200
    assert "No users yet" in r.text


async def test_create_user_then_disable(client: AsyncClient, fakes) -> None:
    r1 = await client.post(
        "/admin/users",
        data={"email": "[email protected]", "name": "Alice", "dni": "12345678A"},
    )
    assert r1.status_code == 201
    assert "[email protected]" in fakes["users"]._users

    user = next(iter(fakes["users"]._users.values()))
    r2 = await client.patch(f"/admin/users/{user.id}/disable")
    assert r2.status_code == 200
    assert fakes["users"].updates == [(user.id, UserStatus.DISABLED)]


async def test_create_user_duplicate_email_returns_409(client: AsyncClient, fakes) -> None:
    await client.post(
        "/admin/users",
        data={"email": "[email protected]", "name": "Alice", "dni": "1"},
    )
    r2 = await client.post(
        "/admin/users",
        data={"email": "[email protected]", "name": "Bob", "dni": "2"},
    )
    assert r2.status_code == 409


async def test_list_apps_renders_empty(client: AsyncClient) -> None:
    r = await client.get("/admin/apps")
    assert r.status_code == 200
    assert "No apps yet" in r.text


async def test_create_assignment_calls_repo(client: AsyncClient, fakes) -> None:
    from uuid import uuid4

    r = await client.post(
        "/admin/assignments",
        data={
            "user_id": str(uuid4()),
            "app_id": "1",
            "profile_id": str(uuid4()),
        },
    )
    assert r.status_code == 201
    assert len(fakes["assignments"].created) == 1


async def test_list_audit_renders_empty(client: AsyncClient) -> None:
    r = await client.get("/admin/audit")
    assert r.status_code == 200
    assert "No events yet" in r.text


async def test_activate_app_renders_confirmation(client: AsyncClient) -> None:
    r = await client.post("/admin/apps/1/activate")
    assert r.status_code == 200
    assert "activada" in r.text


# ---------------------------------------------------------------------------
# Fixture: httpx async client bound to the FastAPI ASGI transport.
# ---------------------------------------------------------------------------


@pytest.fixture
def client(app):
    application, _ = app
    return AsyncClient(transport=ASGITransport(app=application), base_url="http://test")
