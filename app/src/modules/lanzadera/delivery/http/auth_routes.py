# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#543)
"""Auth routes (D-W62-1, AD-W62-2, AD-W62-3).

Three HTTP routes that close the W62 auth flow:

- ``POST /auth/login`` — verify credentials, issue a JWT, create the
  session row.
- ``POST /auth/logout`` — revoke the session identified by the JWT.
- ``GET /auth/me`` — return ``request.state.user_id`` for the caller.

The JWT signing lives in ``container.jwt_signer`` (PR-4 + PR-6) which wraps
the ``JwtSignerPort`` from PR-4. The auth middleware (PR-5) is mounted on
the app in ``main.py`` (also PR-6) and populates ``request.state.user_id``
and ``request.state.session_id`` from the JWT ``sub`` claim before these
routes fire.

Error → HTTP status mapping:
- 401 ``InvalidCredentialsError`` / ``SessionNotFoundError``
- 423 ``AccountLockedError``
- 403 ``AccountNotActiveError``
- 200 / 204 happy path

The ``register_auth_routes`` factory takes the ``LanzaderaContainer`` as a
keyword-only argument so the routes resolve their ports through the
container's properties (``container.jwt_signer``, ``container.sessions``,
``container.audit``) rather than reaching into the di layer's internals.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Annotated

from fastapi import APIRouter, Body, HTTPException, Request, status

from app.src.modules.lanzadera.application.login import login
from app.src.modules.lanzadera.application.logout import logout
from app.src.modules.lanzadera.domain.errors import (
    AccountLockedError,
    AccountNotActiveError,
    InvalidCredentialsError,
    SessionNotFoundError,
)
from app.src.modules.lanzadera.domain.session import LockoutPolicy

if TYPE_CHECKING:
    from app.src.modules.lanzadera.di.container import LanzaderaContainer


def register_auth_routes(router: APIRouter, *, container: LanzaderaContainer) -> None:
    """Mount the auth routes on ``router``.

    The container provides:
    - ``container.jwt_signer`` — ``JwtSignerPort`` for signing the JWT
      (PR-4 + PR-6).
    - ``container.sessions`` — ``SessionRepository`` for create / revoke.
    - ``container.audit`` — ``AuditLog`` for the DA-11 audit rows.
    - ``container.users`` — ``UserRepository`` for the credential lookup.
    - ``container.password_hasher`` — ``PasswordHasher`` for verification.

    Each route handler captures ``container`` via closure so the test app
    (which builds a FastAPI instance with the fakes-backed container) and
    the production ``main.py`` share the exact same wiring.
    """
    from datetime import UTC, datetime

    # W62 PR-6: 24h session TTL comes from the use case constant. The
    # default lockout policy (D38, D39) is created once and reused.
    # Per-user lockout overrides would arrive in a future epic.
    _lockout = LockoutPolicy()

    def _now() -> datetime:
        return datetime.now(UTC)

    @router.post("/auth/login", status_code=status.HTTP_200_OK, tags=["auth"])
    async def login_route(
        req: Request,
        body: Annotated[dict[str, object], Body()],
    ) -> dict[str, str]:
        """Verify credentials, create a session, issue a JWT.

        Body: ``{"email": "...", "password": "..."}``. Returns
        ``{"token": "<jwt>", "session_id": "<uuid>"}`` on success.
        """
        email = body.get("email")
        password = body.get("password")
        if not isinstance(email, str) or not isinstance(password, str):
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="email and password must be strings",
            )
        now = _now()
        try:
            session = await login(
                email,
                password,
                now=now,
                users=container.users,
                sessions=container.sessions,
                password_hasher=container.password_hasher,
                audit=container.audit,
                lockout=_lockout,
            )
        except AccountLockedError as exc:
            raise HTTPException(
                status_code=status.HTTP_423_LOCKED,
                detail="account locked",
            ) from exc
        except AccountNotActiveError as exc:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=f"account not active: {exc}",
            ) from exc
        except InvalidCredentialsError as exc:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="invalid credentials",
            ) from exc

        token = container.jwt_signer.sign(
            {
                "sub": str(session.id),
                "iat": int(now.timestamp()),
                "exp": int(session.expires_at.timestamp()),
            }
        )
        return {"token": token, "session_id": str(session.id)}

    @router.post("/auth/logout", status_code=status.HTTP_204_NO_CONTENT, tags=["auth"])
    async def logout_route(req: Request) -> None:
        """Revoke the session identified by the JWT in ``Authorization``.

        The middleware (PR-5) populates both ``request.state.user_id`` and
        ``request.state.session_id`` from the JWT ``sub`` claim. The logout
        use case (PR-3) revokes the session identified by ``session_id``;
        ``user_id`` flows through as the audit actor. Returns 204 on success
        or 401 if the middleware left the state empty (no token) or the
        session row does not exist.
        """
        user_id = getattr(req.state, "user_id", None)
        session_id = getattr(req.state, "session_id", None)
        if user_id is None or session_id is None:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="authentication required",
            )
        now = _now()
        try:
            await logout(
                session_id,
                now=now,
                sessions=container.sessions,
                audit=container.audit,
                actor_id=user_id,
            )
        except SessionNotFoundError as exc:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="session not found",
            ) from exc
        return None

    @router.get("/auth/me", status_code=status.HTTP_200_OK, tags=["auth"])
    async def me_route(req: Request) -> dict[str, str]:
        """Return ``request.state.user_id`` for the authenticated caller.

        Returns 401 if no token is present (the middleware leaves the
        state empty when the header is missing or the token is invalid).
        """
        user_id = getattr(req.state, "user_id", None)
        if user_id is None:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="authentication required",
            )
        return {"user_id": str(user_id)}


__all__ = ["register_auth_routes"]
