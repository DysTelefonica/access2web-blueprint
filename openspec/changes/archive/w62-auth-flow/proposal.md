# Proposal: W62 — Full auth flow (PR-3 through PR-7)

> **Change**: `w62-auth-flow` · **Project**: `access2web-blueprint` · **Author**: ardelperal · **Date**: 2026-09-01 · **strict_tdd**: true

## Why

W62 cierra la epic Lanzadera MVP con auth real. Reemplaza los stubs `require_global_admin` y `_actor_id` por auth real basada en sesión + JWT. La cadena de **5 PRs encadenados** (PR-3..PR-7) cierra lo que PR-1 (#538, SessionRepository) y PR-2 (#545, login use case) dejaron abierto.

El W58 HTML describe W62 como "Login/session management: auth flow, JWT, logout" pero está desactualizado — W-TEST ya está mergeado y varios sub-piezas ya están en `main` (PR-1 + PR-2). Este change cubre sólo lo que falta.

## What changes

- **`application/logout.py`** (NEW) — caso de uso logout. PR-3.
- **`adapters/crypto/jwt.py`** + **`JwtSignerPort`** (NEW) — utilidad JWT con HS256 + `FakeJwtSigner` para tests. PR-4.
- **`delivery/http/auth_middleware.py`** (NEW) — Bearer token → `request.state.user_id`. PR-5.
- **`delivery/http/auth_routes.py`** (NEW) — `POST /auth/login`, `POST /auth/logout`, `GET /auth/me`. PR-6.
- **`delivery/http/admin.py::require_global_admin`** — implementación real con `request.state.user_id`. PR-6.
- **`delivery/http/admin_routes_{apps,users,misc}.py::_actor_id`** — reemplazar stubs por lectura desde `request.state`. PR-7.
- **`tests/lanzadera/conftest.py`** — retirar `auth_bypass` fixture. PR-7.
- **`tests/lanzadera/delivery/test_admin_routes_integration.py`** — quitar `@pytest.mark.usefixtures("auth_bypass")`. PR-7.

## Decisiones de scope (cerradas)

- **D-W62-1**: HS256 para JWT (sin libs externas de auth — sólo `hmac` + `hashlib` + `secrets` stdlib).
- **D-W62-2**: sesiones persistidas en Postgres con TTL 24h. `Revoked` = `expires_at < now`, no delete (DA-11).
- **D-W62-3**: lockout tras 5 intentos fallidos (D38), duración 1h (D39) — `LockoutPolicy` ya existente.
- **D-W62-4**: logout revoca la sesión actual, no todas las sesiones del usuario. "Logout everywhere" es epic futura.
- **D-W62-5**: `auth_bypass` fixture se retira en PR-7. HR-6 de la skill `lanzadera-testing-strategy` cita el cierre de M02 como trigger.

## Decisiones técnicas (cerradas)

- **JWT library**: cero dependencias externas. La implementación usa `hmac.new(secret, payload, hashlib.sha256)` y `base64.urlsafe_b64encode`/`b64decode`. La "firma" es `HMAC-SHA256(secret, header.payload)`. Header fijo `{"alg":"HS256","typ":"JWT"}`. Payload `{"sub": "<session_id>", "iat": <unix_ts>, "exp": <unix_ts>}`.
- **JWT secret**: leída de `EnvSecretManagerPort.get("JWT_SECRET")`. Tests inyectan un secret fijo. En producción el secret se carga en el bootstrap.
- **Middleware style**: ASGI middleware de Starlette/FastAPI (`BaseHTTPMiddleware`). Lee `Authorization: Bearer <token>`, verifica con el `JwtSignerPort`, popula `request.state.user_id` y `request.state.session_id`. Si el header falta o el token es inválido, deja `request.state.user_id = None` y deja pasar (las rutas destructivas se protegen con `require_global_admin`).
- **Error responses (PR-6)**: HTTP 401 = `Authorization` ausente o token inválido; HTTP 423 = cuenta bloqueada (lockout); HTTP 403 = cuenta no-ACTIVE; HTTP 401 = credenciales inválidas (deliberadamente igual que "no existe" para evitar enumeración). HTTP 401 = sesión no-encontrada en logout.
- **Tests Categoría 2 (login, logout)**: `FakeFixtures` ya inyectado por conftest.py (PR-2). El test del login ya pasa (10 tests, 706 totales). El test del logout debe `await sessions.create(session)` — los fakes son async por DA-1.
- **Tests Categoría 3 (JWT)**: contract test contra `JwtSignerPort`. HS256 round-trip + signature tampering + expiration. La fake `FakeJwtSigner` se usa en tests de Categoría 4.
- **Tests Categoría 4 (middleware, routes)**: conftest provee `LanzaderaContainer` con fakes. Auth bypass NO se usa más (PR-7). Las rutas destructivas se testean con la cookie/session de un admin real (inyectado vía `auth_session` fixture nueva).

## Non-goals

- Cierre de G1 (Postgres e2e en CI), G2 (concurrencia), G3 (atomicidad audit) — esos quedan documentados como gaps en `testing-strategy.md`.
- "Logout everywhere" (revocar todas las sesiones del usuario).
- Refresh tokens / token rotation.
- Rate limiting de login a nivel HTTP (más allá del lockout por usuario).
- 2FA / TOTP.
- Password complexity rules (van en otra epic, F04 probablemente).

## Chain strategy (5 PRs encadenados)

Cada PR parte de la rama anterior. Budget 400 líneas por PR (`CONTRIBUTING.md`). Si una PR lo excede, `size:exception` documentado.

| PR | Título | Pieza | Issues cerradas |
|---|---|---|---|
| PR-3 | `feat/<n>-logout-use-case` | Caso de uso logout + DI wiring + tests Cat 2 | #540 |
| PR-4 | `feat/<n>-jwt-utility` | `JwtSignerPort` + HS256 + `FakeJwtSigner` + tests Cat 3 | #541 |
| PR-5 | `feat/<n>-auth-middleware` | Bearer middleware + `request.state.user_id` + tests Cat 4 | #542 |
| PR-6 | `feat/<n>-auth-routes` | `POST /auth/{login,logout}` + `GET /auth/me` + `require_global_admin` real + tests Cat 4 | #543 |
| PR-7 | `feat/<n>-stub-replacement` | Reemplazar `_actor_id` stubs + retirar `auth_bypass` + auth_session fixture | #544 |

## Lecciones aprendidas (PR-1 + PR-2)

Estas las aplico desde el PR-3 para evitar el ciclo de fixup:

1. **Async fakes**: todo método del fake es `async def`. Tests deben `await` toda llamada. `_seed_session()` en test_logout.py: `await sessions.create(session)`, no `sessions.create(session)`.
2. **TYPE_CHECKING para forward refs**: en login.py uso `from app.src.modules.lanzadera.domain.ports import UserRepository` dentro de `if TYPE_CHECKING:`. Sin esto, mypy pierde el tipo y reporta `attr-defined`.
3. **ruff UP037**: con `from __future__ import annotations`, las quotes alrededor de forward refs son innecesarias. Usar `users: UserRepository` (no `users: "UserRepository"`).
4. **mutation_sites BASELINE ratchet**: cada PR que crezca un archivo cubierto por BASELINE requiere un fixup commit que actualice el sitio. Comando: `python3 scripts/check_mutation_sites.py --emit-baseline` para regenerar.
5. **Container refactor con `_pick`**: ya en PR-2. Cero cambios necesarios al container en PR-3..PR-7 para los repos existentes. PR-6 añade `logout` al factory dict (similar a `login`).
6. **Property type annotations**: `_user_repo` es `object` por `_pick`. Las properties usan `cast(X, self._x_repo)` para mantener el tipo público.
7. **`asyncio_mode = "auto"`** + `-c app/pyproject.toml --rootdir=app` en CI. Tests async no necesitan `@pytest.mark.asyncio`.
8. **Docker daemon intermitente**: rerun con `gh run rerun --failed` o `gh workflow run security.yml --ref <branch>`.

## Acceptance criteria (al cerrar el change)

- [ ] PR-3..PR-7 mergeados en main.
- [ ] `require_global_admin` ya no es stub (verifica `request.state.user_id` contra `GlobalAdminRepository`).
- [ ] `_actor_id` stubs reemplazados por lectura desde `request.state`.
- [ ] Las rutas admin funcionan con JWT válido + devuelven 401 sin token.
- [ ] Tests de integración cubren: login OK, login bad password, login locked, logout, logout no-session, GET /auth/me.
- [ ] Tests de integración retiran `auth_bypass` fixture.
- [ ] `check_test_classification.py --strict` pasa en el repo entero.
- [ ] Coverage global ≥ 69%.
- [ ] `mutation` score semanal no cae.
- [ ] CHANGELOG actualizado con W62 como bloque (ya hay entradas parciales en W54..W61; W62 se añade como "W62 full auth flow").
- [ ] `docs/calidad-de-codigo-y-ci.md` actualizado: conocido PENDING `SessionRepositoryPg` (PR-8 futuro).
- [ ] Issues #540..#544 cerrados al mergear cada PR.
- [ ] Issue #537 (umbrella W62 full auth flow) cerrado al mergear PR-7.

## Risks

- **Riesgo**: la cadena de middleware + JWT + routes tiene varias oportunidades de fallas de typing. Mitigación: aplicar `TYPE_CHECKING` + `cast()` consistentemente desde el primer PR.
- **Riesgo**: `tests/lanzadera/delivery/test_admin_routes_integration.py` ya tiene tests Cat 4 con `auth_bypass`. Reemplazar requiere actualización de cada test que usa el bypass. Mitigación: introducir `auth_session` fixture que popule `request.state.user_id` desde una sesión de fake, y aplicarla donde hoy se aplica `auth_bypass`.
- **Riesgo**: `request.state` es por-request, no global. Cada request HTTP tiene su propio scope. Mitigación: usar el middleware correctamente y verificar en tests que dos requests concurrentes no comparten state.
- **Riesgo**: JWT secret hardcoded en tests. Mitigación: leer de `EnvSecretManager`; tests inyectan un secret fijo vía `EnvSecretManagerAdapter` con override.
- **Riesgo**: el `Hubble` de tests integration Cat 4 con Fakes de Postgres puede crecer el mutation surface. Mitigación: BASELINE actualizada en cada PR per la convención existente.
