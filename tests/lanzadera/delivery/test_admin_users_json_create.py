"""Issue #774 — ``POST /admin/users`` accepts a JSON body.

The pre-fix signature ``*, email: str, name: str, national_id: str``
silently coerced every field into a query parameter and rejected a
real JSON body with ``422 Field required [query, email]`` (the test
below would have asserted 201 against the bug and the assertion
would fail because the route answered 422). With the Pydantic body
model ``CreateUserRequest`` mounted on
``app/src/modules/lanzadera/delivery/http/schemas.py`` the OpenAPI
spec reflects a ``requestBody`` with the schema, and the handler
returns 201 with the serialized user on the happy path.
"""

from __future__ import annotations

from collections.abc import Iterator
from datetime import UTC, datetime
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

import pytest
from fastapi import FastAPI, Request
from fastapi.testclient import TestClient

from app.src.modules.lanzadera.delivery.http.admin_routes_users_json import (
    router as users_json_router,
)
from app.src.modules.lanzadera.delivery.http.auth_middleware import (
    AuthMiddleware,
)
from app.src.modules.lanzadera.di.container import LanzaderaContainer
from app.src.modules.lanzadera.domain.user import User, UserStatus
from tests.lanzadera._fakes import FakeJwtSigner

if TYPE_CHECKING:
    from tests.lanzadera.conftest import FakeFixtures


def _build_container(fakes: FakeFixtures) -> LanzaderaContainer:
    """Wire a container with the FakeFixtures bundle and a stub create_user.

    The handler exercises ``container.create_user``; we replace it
    with a thin shim that builds a domain ``User`` and stores it in
    ``FakeUserRepository.by_id`` so the test sees the create call
    through the existing ``.by_id`` inspection.
    """
    container = LanzaderaContainer(
        session_factory=None,
        secret_manager=fakes.secrets,
        password_hasher=fakes.hasher,
        bootstrap_source=fakes.bootstrap,
        user_repo=fakes.users,
        app_repo=fakes.apps,
        profile_repo=fakes.profiles,
        assignment_repo=fakes.assignments,
        global_admin_repo=fakes.global_admins,
        reset_token_repo=fakes.reset_tokens,
        presence_repo=fakes.presence,
        session_repo=fakes.sessions,
        audit=fakes.audit,
        jwt_signer=FakeJwtSigner(),
    )

    async def _create_user(
        email: str, *, name: str, national_id: str, actor_id: UUID | None = None
    ):
        user = User(
            id=uuid4(),
            email=email,
            name=name,
            dni_encrypted=bytes(f"enc:{national_id}", "utf-8"),
            password_hash=None,
            status=UserStatus.PASSWORD_RESET_REQUIRED,
            failed_attempts=0,
            last_login_at=None,
            created_at=datetime.now(UTC),
            updated_at=datetime.now(UTC),
        )
        fakes.users.add(user)
        return user

    container.create_user = _create_user  # type: ignore[method-assign]
    return container


@pytest.fixture
def admin_user_client(fake_fixtures: FakeFixtures) -> Iterator[TestClient]:
    """Yield a ``TestClient`` with the JSON users router mounted.

    Mirrors what ``app.src.main`` does at boot: a container-mirror
    middleware copies ``app.state.container`` onto ``request.state.container``
    on every request (the JSON routes resolve the container from
    ``request.state.container``, not from ``app.state.container``).
    A bypass middleware then sets ``request.state.user_id`` so the
    destructive ``require_global_admin`` gate short-circuits to the
    repo lookup. The fake's ``FakeGlobalAdminRepository.members`` is
    populated below so the membership check passes.
    """
    from starlette.middleware.base import BaseHTTPMiddleware

    container = _build_container(fake_fixtures)
    admin_id = uuid4()
    fake_fixtures.global_admins.members.add(admin_id)

    app = FastAPI()
    app.state.container = container

    class _ContainerMirror(BaseHTTPMiddleware):
        async def dispatch(self, request, call_next):  # type: ignore[no-untyped-def]
            request.state.container = container
            return await call_next(request)

    app.add_middleware(_ContainerMirror)
    app.include_router(users_json_router)

    @app.middleware("http")
    async def _bypass(request: Request, call_next):  # type: ignore[no-untyped-def]
        request.state.user_id = admin_id
        return await call_next(request)

    # AuthMiddleware still wraps the path so any new dependency wired
    # through it keeps verifying the request lifecycle.
    app.add_middleware(
        AuthMiddleware,
        jwt_signer=container.jwt_signer,
        now=lambda: int(datetime.now(UTC).timestamp()),
    )
    with TestClient(app) as client:
        yield client


def test_create_user_with_json_body_returns_201(
    admin_user_client: TestClient,
) -> None:
    """``POST /admin/users`` with a JSON body returns 201 + the user payload."""
    response = admin_user_client.post(
        "/admin/users",
        headers={"Content-Type": "application/json"},
        json={
            "email": "Issue-774-User@Example.Test",
            "name": "Issue 774",
            "national_id": "ID-774",
        },
    )
    assert response.status_code == 201, response.text
    body = response.json()
    # The schema validates and the handler lower-cases the email
    # before the use case stores it; the response echoes the canonical
    # shape, not the raw client payload.
    assert body["email"] == "issue-774-user@example.test"
    assert body["name"] == "Issue 774"
    assert body["status"] == UserStatus.PASSWORD_RESET_REQUIRED.value
    assert "id" in body and UUID(body["id"])
    assert "created_at" in body


def test_create_user_with_missing_field_returns_422(
    admin_user_client: TestClient,
) -> None:
    """Missing required field → 422 from the Pydantic validator.

    The pre-fix behaviour accepted ``?email=foo`` (and any
    combination of query params) without the body. With the schema
    in place, an incomplete body is rejected by Pydantic before the
    handler runs — pinning this prevents a regression to the old
    signature where the only rejection was the auth gate's 401.
    """
    response = admin_user_client.post(
        "/admin/users",
        headers={"Content-Type": "application/json"},
        json={
            "email": "missing-name@example.test",
            # "name" intentionally absent
            "national_id": "ID-NO-NAME",
        },
    )
    assert response.status_code == 422, response.text
