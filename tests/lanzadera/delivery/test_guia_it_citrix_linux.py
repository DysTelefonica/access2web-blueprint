"""Category 4: HTTP delivery with the real admin guard and container, fake driven ports."""

import re
from datetime import UTC, datetime, timedelta
from pathlib import Path
from uuid import uuid4

import pytest
from fastapi import APIRouter, FastAPI
from fastapi.templating import Jinja2Templates
from httpx import ASGITransport, AsyncClient

from app.src.modules.lanzadera.adapters.crypto.jwt import Hs256JwtSigner
from app.src.modules.lanzadera.delivery.http.admin_routes import register_routes
from app.src.modules.lanzadera.delivery.http.auth_middleware import AuthMiddleware
from app.src.modules.lanzadera.di.container import LanzaderaContainer
from app.src.modules.lanzadera.domain.session import Session
from tests.lanzadera._fakes import (
    FakeAppRepository,
    FakeAssignmentRepository,
    FakeAuditLog,
    FakeBootstrapAdminSource,
    FakeGlobalAdminRepository,
    FakePasswordHasher,
    FakeProfileRepository,
    FakeResetTokenRepository,
    FakeSecretManager,
    FakeSessionRepository,
    FakeUserRepository,
)
from tests.lanzadera._presence_fakes import FakePresenceRepository

ROOT = Path(__file__).resolve().parents[3]
TEMPLATES = ROOT / "app/src/modules/lanzadera/delivery/http/templates"
GUIDE = "/admin/guia-it-citrix-linux"


@pytest.fixture
async def guide_client():
    signer = Hs256JwtSigner(b"test-guide-signing-key-32-bytes!!")
    sessions = FakeSessionRepository()
    admins = FakeGlobalAdminRepository()
    container = LanzaderaContainer(
        secret_manager=FakeSecretManager(),
        password_hasher=FakePasswordHasher(),
        bootstrap_source=FakeBootstrapAdminSource(),
        user_repo=FakeUserRepository(),
        app_repo=FakeAppRepository(),
        profile_repo=FakeProfileRepository(),
        assignment_repo=FakeAssignmentRepository(),
        global_admin_repo=admins,
        reset_token_repo=FakeResetTokenRepository(),
        presence_repo=FakePresenceRepository(),
        audit=FakeAuditLog(),
        session_repo=sessions,
        jwt_signer=signer,
    )
    app = FastAPI()
    app.state.container = container
    app.add_middleware(AuthMiddleware, jwt_signer=signer)
    router = APIRouter(prefix="/admin")
    register_routes(
        router, templates=Jinja2Templates(directory=str(TEMPLATES)), container=container
    )
    app.include_router(router)
    now = datetime.now(UTC)
    tokens = {}
    for kind in ("admin", "user"):
        user_id, session_id = uuid4(), uuid4()
        sessions.sessions[session_id] = Session(
            id=session_id, user_id=user_id, created_at=now, expires_at=now + timedelta(hours=1)
        )
        if kind == "admin":
            admins.members.add(user_id)
        tokens[kind] = signer.sign(
            {
                "sub": str(session_id),
                "iat": int(now.timestamp()),
                "exp": int(now.timestamp()) + 3600,
            }
        )
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        yield client, tokens, container


@pytest.mark.parametrize("identity, expected", [(None, 401), ("user", 403), ("admin", 200)])
async def test_guide_uses_existing_admin_authorization(guide_client, identity, expected):
    client, tokens, container = guide_client
    headers = {"Authorization": f"Bearer {tokens[identity]}"} if identity else {}
    before = set(container.global_admin_repo.members)
    response = await client.get(GUIDE, headers=headers)
    assert response.status_code == expected
    assert container.global_admin_repo.members == before
    if expected != 200:
        assert "APLICACIONES_DYS" not in response.text
        assert "Aplicaciones DYSN" not in response.text
    else:
        assert response.headers["cache-control"] == "no-store"
        for section in (
            "escenarios",
            "red",
            "windows",
            "linux",
            "validacion",
            "entrega",
            "fuentes",
        ):
            assert f'id="{section}"' in response.text
        assert "%APPDATA%" in response.text
        assert 'id="guia-export"' in response.text


async def test_guide_glossary_defines_every_linked_term(guide_client):
    client, tokens, _ = guide_client
    response = await client.get(GUIDE, headers={"Authorization": f"Bearer {tokens['admin']}"})
    assert response.status_code == 200
    html = response.text
    assert 'id="glosario"' in html
    assert 'href="#glosario"' in html
    linked = set(re.findall(r'href="#(g-[a-z0-9-]+)"', html))
    defined = set(re.findall(r'id="(g-[a-z0-9-]+)"', html))
    assert "g-bitness" in linked
    assert len(linked) >= 30
    assert linked <= defined


async def test_admin_navigation_and_audit_keep_existing_behavior(guide_client):
    client, tokens, _ = guide_client
    response = await client.get(
        "/admin/audit", headers={"Authorization": f"Bearer {tokens['admin']}"}
    )
    assert response.status_code == 200
    for route in ("/admin/users", "/admin/apps", "/admin/assignments", "/admin/audit", GUIDE):
        assert f'href="{route}"' in response.text
    anonymous = await client.get("/admin/audit")
    assert anonymous.status_code == 401
    regular = await client.get(
        "/admin/audit", headers={"Authorization": f"Bearer {tokens['user']}"}
    )
    assert regular.status_code == 403
