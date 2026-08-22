"""Contract tests for the auth HTTP delivery (DL1, issue #54).

The four endpoints (login GET/POST, logout POST, reset GET/POST) are
tested with ``httpx.AsyncClient`` + ``ASGITransport`` for in-process
FastAPI testing. The use-case ports are fakes; the reset flow uses a
``FakeGlobalAdminRepository`` because the D90 service requires a global
admin to exist (DA-6).
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Iterator
from uuid import UUID

import pytest
from fastapi import FastAPI
from fastapi.templating import Jinja2Templates
from httpx import ASGITransport, AsyncClient

from app.src.modules.lanzadera.delivery.http.auth import build_router
from app.src.modules.lanzadera.domain.reset_token import ResetToken
from app.src.modules.lanzadera.domain.user import User, UserStatus
from app.src.modules.lanzadera.domain.ports.notification_delivery import (
    NotificationDeliveryPort,
)


# ---------------------------------------------------------------------------
# Fakes.
# ---------------------------------------------------------------------------


class FakeUserRepository:
    def __init__(self) -> None:
        self._users: dict[str, User] = {}

    async def get_by_email(self, email: str) -> User | None:
        return self._users.get(email.lower())

    async def list_all(self) -> list[User]:
        return list(self._users.values())


class FakeResetTokenRepository:
    def __init__(self) -> None:
        self.tokens: list[ResetToken] = []

    async def insert(
        self, user_id: UUID, token_hash: str, expires_at: datetime
    ) -> ResetToken:
        token = ResetToken(
            id=UUID(int=0xDEAD),
            user_id=user_id,
            token_hash=token_hash,
            expires_at=expires_at,
            consumed_at=None,
            superseded_at=None,
            created_at=datetime.now(UTC),
        )
        self.tokens.append(token)
        return token

    async def mark_superseded(self, user_id: UUID, at: datetime) -> None:
        for t in self.tokens:
            if t.user_id == user_id and t.superseded_at is None:
                t.superseded_at = at
                return


class FakeAuditLogPort:
    async def append(self, event) -> None:  # noqa: ARG002
        return None


class FakeNotificationDelivery(NotificationDeliveryPort):
    async def send(self, to: str, subject: str, body: str) -> None:  # noqa: ARG002
        return None


class FakeGlobalAdminRepository:
    async def is_global_admin(self, user_id: UUID) -> bool:
        return True

    async def there_is_any(self) -> bool:
        return True


# ---------------------------------------------------------------------------
# Fixture.
# ---------------------------------------------------------------------------


@pytest.fixture
def fakes():
    return {
        "users": FakeUserRepository(),
        "reset_tokens": FakeResetTokenRepository(),
        "notifications": FakeNotificationDelivery(),
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
            reset_tokens=fakes["reset_tokens"],
            notifications=fakes["notifications"],
            global_admins=fakes["admins"],
        )
    )
    return application, fakes


# ---------------------------------------------------------------------------
# Tests.
# ---------------------------------------------------------------------------


async def test_get_login_renders_form(client: AsyncClient) -> None:
    r = await client.get("/login")
    assert r.status_code == 200
    assert "Iniciar" in r.text or "Entrar" in r.text


async def test_get_reset_renders_form(client: AsyncClient) -> None:
    r = await client.get("/reset")
    assert r.status_code == 200
    assert "Recuperar" in r.text or "contraseña" in r.text


async def test_post_login_redirects_to_admin_on_known_email(client: AsyncClient) -> None:
    user = User(
        id=UUID(int=0xCAFE),
        email="[email protected]",
        name="Alice",
        dni_encrypted=b"x",
        password_hash="hash",
        status=UserStatus.ACTIVE,
        failed_attempts=0,
        last_login_at=None,
        created_at=datetime.now(UTC),
        updated_at=datetime.now(UTC),
    )
    r = await client.post(
        "/login",
        data={"email": "[email protected]", "password": "any"},
    )
    assert r.status_code == 303
    assert r.headers["location"] == "/admin"


async def test_post_login_redirects_to_login_on_unknown_email(client: AsyncClient) -> None:
    r = await client.post(
        "/login",
        data={"email": "[email protected]", "password": "any"},
    )
    assert r.status_code == 303
    assert r.headers["location"] == "/login?error=invalid"


async def test_post_logout_redirects_to_login(client: AsyncClient) -> None:
    r = await client.post("/logout")
    assert r.status_code == 303
    assert r.headers["location"] == "/login"


async def test_post_reset_invokes_issue_reset_token(client: AsyncClient, fakes) -> None:
    user = User(
        id=UUID(int=0xCAFE),
        email="[email protected]",
        name="Alice",
        dni_encrypted=b"x",
        password_hash=None,
        status=UserStatus.PASSWORD_RESET_REQUIRED,
        failed_attempts=0,
        last_login_at=None,
        created_at=datetime.now(UTC),
        updated_at=datetime.now(UTC),
    )
    fakes["users"]._users["[email protected]"] = user

    r = await client.post("/reset", data={"email": "[email protected]"})
    assert r.status_code == 303
    assert r.headers["location"] == "/login?reset=sent"
    assert len(fakes["reset_tokens"].tokens) == 1


# ---------------------------------------------------------------------------
# Client fixture.
# ---------------------------------------------------------------------------


@pytest.fixture
def client(app):
    application, _ = app
    return AsyncClient(transport=ASGITransport(app=application), base_url="http://test")