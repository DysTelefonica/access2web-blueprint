# Spec: Auth routes (PR-6)

> Capability del change `w62-auth-flow`. Tres rutas HTTP que cierran el flujo de auth: `POST /auth/login`, `POST /auth/logout`, `GET /auth/me`.

## Purpose

Las rutas admin actuales usan `auth_bypass` fixture en tests. PR-6 reemplaza el stub `require_global_admin` con una implementación real que verifica `request.state.user_id` (poblado por el middleware de PR-5) y consulta `GlobalAdminRepository` para confirmar el rol.

Además añade las tres rutas de auth que faltan para que el flujo end-to-end funcione.

## Requirements

### Requirement: POST /auth/login

The system SHALL accept ``{"email": "...", "password": "..."}`` and
return ``{"token": "<jwt>", "session_id": "<uuid>"}`` on success.

#### Scenario: Successful login

- GIVEN a request with valid credentials
- WHEN ``POST /auth/login`` is invoked
- THEN the response is 200 with body ``{"token": "<jwt>", "session_id": "<session_id>"}``
- AND the JWT's ``sub`` claim equals the session_id

#### Scenario: Invalid credentials

- GIVEN a request with wrong password or unknown email
- WHEN ``POST /auth/login`` is invoked
- THEN the response is 401 with body ``{"detail": "Invalid credentials"}``
- AND the JWT is NOT issued

#### Scenario: Account locked

- GIVEN a user whose ``failed_attempts >= LockoutPolicy.threshold`` and last attempt within the lockout window
- WHEN ``POST /auth/login`` is invoked
- THEN the response is 423 with body ``{"detail": "Account locked"}``

### Requirement: POST /auth/logout

The system SHALL accept ``Authorization: Bearer <jwt>`` and revoke the
session matching the JWT's ``sub`` claim.

#### Scenario: Successful logout

- GIVEN a request with a valid Bearer token
- WHEN ``POST /auth/logout`` is invoked
- THEN the response is 204
- AND the session identified by the JWT's ``sub`` claim is revoked (expires_at = now)

#### Scenario: Logout with no token

- GIVEN a request without ``Authorization``
- WHEN ``POST /auth/logout`` is invoked
- THEN the response is 401 with body ``{"detail": "Authentication required"}``

### Requirement: GET /auth/me

The system SHALL return ``{"user_id": "<uuid>"}`` for any authenticated
request.

#### Scenario: Authenticated request

- GIVEN a request with a valid Bearer token
- WHEN ``GET /auth/me`` is invoked
- THEN the response is 200 with body ``{"user_id": "<session_id>"}``

#### Scenario: Unauthenticated request

- GIVEN a request without ``Authorization``
- WHEN ``GET /auth/me`` is invoked
- THEN the response is 401

### Requirement: require_global_admin real implementation

The system SHALL replace the D91 placeholder
``admin.require_global_admin()`` with an implementation that reads
``request.state.user_id`` and raises ``HTTPException(403)`` when the
caller is not a global admin.

```python
async def require_global_admin(request: Request) -> None:
    user_id = getattr(request.state, "user_id", None)
    if user_id is None:
        raise HTTPException(401, "Authentication required")
    container = request.app.state.container
    if not await container.global_admin_repo.is_global_admin(user_id):
        raise HTTPException(403, "Forbidden")
```

## File-surface contract

| File | Action | Notes |
|---|---|---|
| `app/src/modules/lanzadera/delivery/http/auth_routes.py` | NEW | `register_auth_routes(router, *, container, jwt_signer)` mounts the three routes. |
| `app/src/modules/lanzadera/delivery/http/admin.py` | MODIFY | replace stub `require_global_admin` with the real implementation. |
| `app/src/modules/lanzadera/di/container.py` | MODIFY | add `app.pytest_plugin.coverage_gate` injection point: `self._jwt_signer = jwt_signer or Hs256JwtSigner(secret_from_env)`. |
| `app/src/modules/lanzadera/delivery/http/main.py` (or similar entry point) | MODIFY | wire `AuthMiddleware` + `register_auth_routes` into the FastAPI app. **NOTE**: this is `app/src/main.py`, not a new file. |
| `tests/lanzadera/delivery/test_auth_routes.py` | NEW | Categoría 4 integration — 8 test cases. |
| `tests/lanzadera/delivery/test_admin_routes_integration.py` | MODIFY | keep `auth_bypass` for now; PR-7 retires it. |

## Container wiring (in PR-6)

```python
# di/container.py (new constructor parameter)
jwt_signer: JwtSignerPort | None = None,

# inside __init__ (next to the session_repo pattern)
self._jwt_signer = (
    jwt_signer if jwt_signer is not None else Hs256JwtSigner(
        secret=EnvSecretManagerAdapter().get("JWT_SECRET").encode("utf-8")
    )
)

# new factory entry (next to login/logout)
"sign_jwt": functools.partial(
    lambda payload, signer: signer.sign(payload),
    signer=self._jwt_signer,
),
```

## main.py wiring (in PR-6)

```python
# app/src/main.py (existing file)
from starlette.middleware import Middleware
from app.src.modules.lanzadera.delivery.http.auth_middleware import AuthMiddleware
from app.src.modules.lanzadera.delivery.http.auth_routes import register_auth_routes

app.add_middleware(AuthMiddleware, jwt_signer=container.jwt_signer)
register_auth_routes(auth_router, container=container)
```

(Exact location depends on the existing main.py structure; PR-6 will adapt.)

## Tests Cat 4 — `tests/lanzadera/delivery/test_auth_routes.py`

8 cases covering all the scenarios above. Each test:
- Builds a `FakeContainer` (or uses the real `LanzaderaContainer` with fakes).
- Calls `httpx.AsyncClient` against an ASGI app that mounts the routes.
- Asserts status code + body.

## Decisiones cerradas

- **Body schema**: snake_case (`session_id`, `user_id`).
- **Error body**: `{"detail": "..."}` (FastAPI default).
- **Content-Type**: `application/json` for login/me; logout is 204 (no body).
- **`require_global_admin`** sigue siendo callable desde las rutas admin (no se cambia la firma). El cambio es interno: pasa de ser un stub a verificar contra `GlobalAdminRepository`.

## BASELINE updates

- `auth_routes.py`: nuevo archivo. Si > 100 mutation sites, agregar al BASELINE.
- `main.py`: si crece > 100, BASELINE update.
- `container.py`: +1 entry en factory + 1 nuevo param. Crece ~5-10 sitios. BASELINE update si supera.
- `admin.py`: el require_global_admin real reemplaza el stub. Mismo número de sites.

## Verificación

```bash
# Sólo los nuevos tests
/usr/bin/python3.12 -m pytest -c app/pyproject.toml --rootdir=app tests/lanzadera/delivery/test_auth_routes.py -v
# → 8 passed

# Suite completa
/usr/bin/python3.12 -m pytest -c app/pyproject.toml --rootdir=app --cov=app --cov-report=term
# → 725 + 8 = 733 passed

# Smoke contra main.py (smoke test verifica que el app monta)
# → covered by tests/lanzadera/delivery/test_auth_routes.py
```
