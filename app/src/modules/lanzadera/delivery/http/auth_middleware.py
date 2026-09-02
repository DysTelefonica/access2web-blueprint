# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#542)
"""AuthMiddleware (D-W62-1, AD-W62-2).

Starlette ASGI middleware that reads the ``Authorization: Bearer <token>``
header on every request, verifies the token with ``JwtSignerPort``, and
populates ``request.state.user_id`` + ``request.state.session_id``.

Silent-failure contract (AD-W62-2): a missing or invalid token does NOT
raise. The middleware leaves ``request.state.user_id = None`` and lets
the request through. The destructive routes' ``require_global_admin``
gate (PR-6) is what returns 401/403 when the state is empty. Public
routes (``/auth/login``, ``/health``) work without auth and the
middleware does not block them.

Per-request state isolation (D-W62-2): ``request.state`` is per-request
in Starlette/FastAPI. Two concurrent requests cannot share state. The
test ``test_middleware_is_per_request`` pins this property.
"""

from __future__ import annotations

import time
from collections.abc import Awaitable, Callable
from uuid import UUID

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import Response
from starlette.types import ASGIApp

from app.src.modules.lanzadera.domain.errors import (
    ExpiredTokenError,
    InvalidTokenError,
)
from app.src.modules.lanzadera.domain.ports.jwt_signer import JwtSignerPort


class AuthMiddleware(BaseHTTPMiddleware):
    """ASGI middleware that decodes the Bearer token (D-W62-1) and
    populates ``request.state.user_id`` + ``request.state.session_id``.

    AD-W62-2 silent-failure contract:
      - Missing ``Authorization`` header → ``request.state.user_id = None``.
      - ``Authorization: Basic ...`` (non-Bearer) → ``request.state.user_id = None``.
      - Malformed / unparseable token → ``request.state.user_id = None``.
      - Expired token → ``request.state.user_id = None``.
      - Signature does not verify → ``request.state.user_id = None``.

    The middleware never raises; the routes' ``require_global_admin``
    is what turns ``state.user_id = None`` into HTTP 401/403.

    Two concurrent requests cannot share state because
    ``request.state`` is per-request in Starlette/FastAPI.
    """

    def __init__(
        self,
        app: ASGIApp,
        *,
        jwt_signer: JwtSignerPort,
        now: Callable[[], int] | None = None,
    ) -> None:
        super().__init__(app)
        self._signer = jwt_signer
        # Inject a clock so tests can pin ``now`` deterministically;
        # production defaults to ``int(time.time())``.
        self._now = now if now is not None else (lambda: int(time.time()))

    async def dispatch(
        self, request: Request, call_next: Callable[[Request], Awaitable[Response]]
    ) -> Response:
        # Default state: no auth. Public routes work; protected routes
        # gate on ``request.state.user_id`` and return 401 when None.
        request.state.user_id = None
        request.state.session_id = None

        auth = request.headers.get("authorization")
        if auth is not None and auth.lower().startswith("bearer "):
            token = auth[7:].strip()
            if token:
                try:
                    payload = self._signer.verify(token, now=self._now())
                except (InvalidTokenError, ExpiredTokenError, ValueError, TypeError):
                    # Silent failure — the auth middleware is not the
                    # gate. The route's ``require_global_admin`` returns
                    # 401/403 if the state is empty.
                    payload = None
                if payload is not None:
                    sub = payload.get("sub")
                    if sub is not None:
                        try:
                            session_id = self._coerce_session_id(sub)
                        except (ValueError, TypeError, AttributeError):
                            session_id = None
                        if session_id is not None:
                            request.state.session_id = session_id
                            request.state.user_id = session_id

        return await call_next(request)

    @staticmethod
    def _coerce_session_id(value: object) -> UUID | None:
        """Coerce the JWT ``sub`` claim into a UUID; return None if invalid.

        The JWT carries the session id as the ``sub`` claim. A token whose
        ``sub`` is not a parseable UUID is malformed; we silently skip
        the request state rather than crashing.
        """
        if isinstance(value, UUID):
            return value
        if isinstance(value, str):
            return UUID(value)
        return None


__all__ = ["AuthMiddleware"]
