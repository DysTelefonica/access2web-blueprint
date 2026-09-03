# Spec: Auth middleware (PR-5)

> Capability del change `w62-auth-flow`. ASGI middleware que extrae el `Authorization: Bearer <token>` y popula `request.state.user_id`.

## Purpose

Las rutas destructivas (POST /admin/users, PATCH /admin/apps, DELETE /admin/apps/{id}) necesitan saber quién es el actor. El middleware es el único responsable de extraer esa información del request y exponerla vía `request.state`.

Decisión: middleware Starlette/FastAPI (`BaseHTTPMiddleware`). Lee `Authorization` header, verifica con `JwtSignerPort`, popula state. **Si falta el header o el token es inválido, deja `request.state.user_id = None` y deja pasar el request** — las rutas destructivas usan `require_global_admin()` que retorna 401 si el state es None.

## Requirements

### Requirement: Extract Bearer token from Authorization header

The system SHALL read the ``Authorization`` header on every request
and extract the JWT after the ``Bearer `` prefix (case-insensitive
on the scheme; case-sensitive on the token).

#### Scenario: Valid Bearer token

- GIVEN a request with header ``Authorization: Bearer <valid_jwt>``
- WHEN the middleware runs
- THEN ``request.state.user_id`` is set to the JWT's ``sub`` claim
- AND ``request.state.session_id`` is set to the JWT's ``sub`` claim (UUID)
- AND the request proceeds

#### Scenario: Missing Authorization header

- GIVEN a request with no ``Authorization`` header
- WHEN the middleware runs
- THEN ``request.state.user_id`` is ``None``
- AND the request proceeds (the destructive route returns 401 via ``require_global_admin``)

#### Scenario: Non-Bearer scheme

- GIVEN a request with header ``Authorization: Basic abcdef``
- WHEN the middleware runs
- THEN ``request.state.user_id`` is ``None`` (scheme mismatch — ignore)
- AND the request proceeds

#### Scenario: Malformed Bearer token

- GIVEN a request with header ``Authorization: Bearer not.a.jwt``
- WHEN the middleware runs
- THEN ``request.state.user_id`` is ``None``
- AND the request proceeds (the JWT signer raises; middleware catches)

#### Scenario: Expired Bearer token

- GIVEN a request with a JWT whose ``exp`` is in the past
- WHEN the middleware runs
- THEN ``request.state.user_id`` is ``None``
- AND the request proceeds

### Requirement: Use injected JwtSignerPort

The system SHALL receive the ``JwtSignerPort`` instance via constructor
injection (per W-TEST DI pattern). Tests inject ``FakeJwtSigner``;
production wires ``Hs256JwtSigner`` (PR-4) with the real secret.

### Requirement: Per-request state isolation

The system SHALL populate ``request.state`` (not module-level globals).
Two concurrent requests on the same app MUST NOT share state.

## File-surface contract

| File | Action | Notes |
|---|---|---|
| `app/src/modules/lanzadera/delivery/http/auth_middleware.py` | NEW | `class AuthMiddleware(BaseHTTPMiddleware)` with `__init__(app, jwt_signer)`. |
| `tests/lanzadera/delivery/test_auth_middleware.py` | NEW | Categoría 4 integration — 6 test cases via httpx.AsyncClient + ASGITransport. |

## Container integration (NO)

Este PR NO toca `LanzaderaContainer`. La integración del middleware en el `main.py` viene en PR-6. Aquí sólo se construye el middleware + tests Cat 4.

## Implementation reference

```python
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request

class AuthMiddleware(BaseHTTPMiddleware):
    def __init__(self, app, *, jwt_signer: JwtSignerPort, now: Callable[[], int] | None = None) -> None:
        super().__init__(app)
        self._signer = jwt_signer
        self._now = now or (lambda: int(time.time()))

    async def dispatch(self, request: Request, call_next):
        request.state.user_id = None
        request.state.session_id = None
        auth = request.headers.get("authorization")
        if auth and auth.lower().startswith("bearer "):
            token = auth[7:].strip()
            try:
                payload = self._signer.verify(token, now=self._now())
                sub = payload.get("sub")
                if sub is not None:
                    request.state.session_id = UUID(sub)
                    request.state.user_id = request.state.session_id
            except (InvalidTokenError, ExpiredTokenError, ValueError):
                pass  # leave state None; routes gate on require_global_admin
        return await call_next(request)
```

Note: the JWT's `sub` claim carries the `session_id`, which doubles as the actor identifier for now (D-W62-4: revoke one session at a time). PR-7 may later split actor vs session identity, but for now `request.state.user_id = request.state.session_id` is correct.

## Tests Cat 4 — `tests/lanzadera/delivery/test_auth_middleware.py`

Tests use `httpx.AsyncClient` + `ASGITransport` against a minimal FastAPI app that mounts the middleware. The app has one route (`/probe`) that returns `{"user_id": str(request.state.user_id) if request.state.user_id else None}`.

1. `test_middleware_passes_through_without_authorization` — no header → 200, user_id is None.
2. `test_middleware_extracts_user_id_from_valid_bearer` — valid JWT → 200, user_id matches `sub` claim.
3. `test_middleware_sets_user_id_to_none_for_malformed_token` — bad token → 200, user_id is None.
4. `test_middleware_sets_user_id_to_none_for_expired_token` — expired JWT → 200, user_id is None.
5. `test_middleware_ignores_non_bearer_scheme` — `Basic abcdef` → 200, user_id is None.
6. `test_middleware_is_per_request` — two concurrent requests with different JWTs both see their own user_id.

## Decisiones cerradas

- **Middleware ASGI** (no WSGI): Starlette/FastAPI corren ASGI nativamente.
- **`request.state` per-request**: Starlette/FastAPI garantiza esto por-request.
- **Silent failure**: middleware NO raise; deja state None. Las rutas destructivas usan `require_global_admin()`.
- **`user_id == session_id`**: por ahora, hasta que llegue un actor_id separado en una epic futura.

## BASELINE updates

`auth_middleware.py` debería quedar bajo 100 mutation sites (código pequeño).

## Verificación

```bash
# Sólo los nuevos tests
/usr/bin/python3.12 -m pytest -c app/pyproject.toml --rootdir=app tests/lanzadera/delivery/test_auth_middleware.py -v
# → 6 passed

# Suite completa
/usr/bin/python3.12 -m pytest -c app/pyproject.toml --rootdir=app --cov=app --cov-report=term
# → 719 + 6 = 725 passed

# Gates
ruff format --check --config app/pyproject.toml .
ruff check --config app/pyproject.toml .
/usr/bin/python3.12 -m mypy --explicit-package-bases app/
python3 scripts/check_test_classification.py --root tests/lanzadera
```
