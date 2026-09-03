# Tasks: W62 auth flow (PR-3..PR-7)

> Change: `w62-auth-flow` · Project: `access2web-blueprint` · strict_tdd: true · pytest: `pytest --cov=app --cov-fail-under=85`

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Chain strategy | chained PRs (PR-3..PR-7 each on top of the previous) |
| 400-line budget risk | Medium (PR-6 likely needs size:exception ~600 lines) |
| Decision needed before apply | Yes — the SDD must be approved before any PR |
| Chained PRs recommended | Yes |

## Phase 1 — Logout use case (PR-3)

Issue: #540

### T-PR3-1: `application/logout.py` (NEW)

`async def logout(session_id: UUID, *, now: datetime, sessions: SessionRepository, audit: AuditLog, actor_id: UUID | None = None) -> None`.

- 5 paths: success / not-found / double-revoke (idempotent) / None session_id / session-disappeared-during-call.
- Revoke = `sessions.revoke(session_id, now)`.
- Audit row in every path (DA-11).

### T-PR3-2: `domain/errors.py` (MODIFY)

Add `class SessionNotFoundError(Exception)`.

### T-PR3-3: `di/container.py` (MODIFY)

Add `async def logout(self, session_id: UUID, *, actor_id: UUID | None = None) -> None` calling `self._use_cases["logout"](...)`.

### T-PR3-4: `di/use_cases.py` (MODIFY)

Add `from app.src.modules.lanzadera.application.logout import logout`. Add `"logout": functools.partial(logout, sessions=session_repo, audit=audit)` to the factory dict.

### T-PR3-5: `tests/lanzadera/application/test_logout.py` (NEW)

7 Cat 2 test cases (see spec). Use `_seed_session()` helper that **awaits** `sessions.create(session)` (async fake — learn from PR-2 worktree failure).

### T-PR3-6: Verify

- `pytest tests/lanzadera/application/test_logout.py -v` → 7 passed
- `pytest --cov=app --cov-report=term` → coverage ≥ 82%, 713 total
- `ruff format && ruff check && mypy && check_test_classification` clean
- If `application/logout.py` exceeds 100 mutation sites, update `scripts/check_mutation_sites.py` BASELINE.

### T-PR3-7: Commit + push + PR

```bash
git checkout -b feat/540-logout-use-case origin/a8009bd
# ... write files ...
git commit -m "feat(lanzadera): add logout use case + container wiring (W62 PR-3)"
git push -u origin feat/540-logout-use-case
gh pr create --base main --title "feat(lanzadera): add logout use case + container wiring (W62 PR-3)"
# Wait for CI green; rerun infra failures (trivy, gitleaks) up to 3x.
gh pr merge --squash --delete-branch=false
gh issue close 540 --comment "..."
```

## Phase 2 — JWT utility (PR-4)

Issue: #541

### T-PR4-1: `domain/ports/jwt_signer.py` (NEW)

`class JwtSignerPort(Protocol)` with `sign(payload: dict[str, Any]) -> str` and `verify(token: str, *, now: int) -> dict[str, Any]`. Add `JwtSigner = JwtSignerPort` alias.

### T-PR4-2: `domain/ports/_imports.py` (MODIFY)

Add `Any` to the prelude if not present.

### T-PR4-3: `adapters/crypto/jwt.py` (NEW)

`class Hs256JwtSigner`:
- `__init__(secret: bytes)` — validate `len(secret) >= 32`.
- `sign(payload) -> str` — HS256 JWT, header `{"alg":"HS256","typ":"JWT"}`.
- `verify(token, *, now: int) -> dict[str, Any]` — verify signature + `exp`.

### T-PR4-4: `domain/errors.py` (MODIFY)

Add `class InvalidTokenError(Exception)` and `class ExpiredTokenError(Exception)`.

### T-PR4-5: `tests/lanzadera/_fakes.py` (MODIFY)

Add `class FakeJwtSigner`:
- `__init__(secret: bytes)`.
- `sign_calls: list[dict[str, Any]]` and `verify_calls: list[tuple[str, int]]` for introspection.
- `sign(payload) -> str` — uses `Hs256JwtSigner` internally (so fake behaves identically) but tracks calls.
- `verify(token, *, now: int) -> dict[str, Any]` — delegates.

### T-PR4-6: `tests/lanzadera/adapters/crypto/test_jwt_signer.py` (NEW)

6 Cat 3 contract tests (see spec). Use `pytest.raises(InvalidTokenError)` / `ExpiredTokenError`.

### T-PR4-7: Verify + commit + push + PR + merge

Standard flow. Issue #541 close after merge.

## Phase 3 — Auth middleware (PR-5)

Issue: #542

### T-PR5-1: `delivery/http/auth_middleware.py` (NEW)

`class AuthMiddleware(BaseHTTPMiddleware)`:
- `__init__(app, *, jwt_signer: JwtSignerPort, now: Callable[[], int] | None = None)`.
- `async def dispatch(request, call_next)`:
  - Default `request.state.user_id = None`, `request.state.session_id = None`.
  - If `Authorization` header starts with `Bearer `, extract token, verify, populate state.
  - Silent on `InvalidTokenError` / `ExpiredTokenError`.
  - Call `await call_next(request)`.

### T-PR5-2: `tests/lanzadera/delivery/test_auth_middleware.py` (NEW)

6 Cat 4 integration tests via `httpx.AsyncClient` + `ASGITransport`. Probe route returns `request.state.user_id`.

### T-PR5-3: Verify + commit + push + PR + merge

Standard. Issue #542 close.

## Phase 4 — Auth routes (PR-6)

Issue: #543

### T-PR6-1: `delivery/http/auth_routes.py` (NEW)

`def register_auth_routes(router: APIRouter, *, container: LanzaderaContainer, jwt_signer: JwtSignerPort)`:
- `POST /auth/login` — body `{"email", "password"}` → 200 `{"token", "session_id"}` or 401/423.
- `POST /auth/logout` — Bearer token → 204.
- `GET /auth/me` — Bearer token → 200 `{"user_id"}`.

### T-PR6-2: `delivery/http/admin.py` (MODIFY)

Replace stub `require_global_admin()` with the real implementation:

```python
async def require_global_admin(request: Request) -> None:
    user_id = getattr(request.state, "user_id", None)
    if user_id is None:
        raise HTTPException(401, "Authentication required")
    container = request.app.state.container
    if not await container.global_admin_repo.is_global_admin(user_id):
        raise HTTPException(403, "Forbidden")
```

### T-PR6-3: `di/container.py` (MODIFY)

Add `jwt_signer: JwtSignerPort | None = None` constructor parameter. Build default `Hs256JwtSigner(secret=EnvSecretManagerAdapter().get("JWT_SECRET").encode("utf-8"))` when None.

Add `self.jwt_signer` property.

### T-PR6-4: `di/use_cases.py` (MODIFY)

No new use case; auth_routes access `container.login` and `container.logout` directly (they are async methods on the container).

### T-PR6-5: `app/src/main.py` (MODIFY)

Wire `AuthMiddleware` + `register_auth_routes` into the FastAPI app. **NOTE**: the actual wiring depends on existing main.py structure; PR-6 adapts.

### T-PR6-6: `tests/lanzadera/delivery/test_auth_routes.py` (NEW)

8 Cat 4 integration tests via httpx.AsyncClient (see spec).

### T-PR6-7: Verify + commit + push + PR + merge

Standard. Likely needs size:exception (PR-6 is the largest). Issue #543 close.

## Phase 5 — Stub replacement (PR-7)

Issue: #544

### T-PR7-1: `delivery/http/admin_routes_{apps,users,misc}.py` (MODIFY)

Replace `_actor_id(request)` stub:

```python
def _actor_id(request: Request) -> UUID | None:
    return getattr(request.state, "user_id", None)
```

### T-PR7-2: `tests/lanzadera/conftest.py` (MODIFY)

Remove `auth_bypass` fixture definition. Add `auth_session` fixture that:
- Seeds a global admin in `FakeGlobalAdminRepository.grant(admin_id)`.
- Seeds a session row in `FakeSessionRepository.sessions[session_id]`.
- Yields `{"admin_id": admin_id, "session_id": session_id}` for tests that need to call the route with a specific session id.

### T-PR7-3: `tests/lanzadera/delivery/test_admin_routes_integration.py` (MODIFY)

For each test currently using `@pytest.mark.usefixtures("auth_bypass")`:
- Replace with `@pytest.mark.usefixtures("auth_session")`.
- Verify the response is the expected success (not 401/403).

The admin routes test app must be updated to use a test-only middleware that reads `X-Test-Session-Id` and populates `request.state.user_id` — this avoids requiring JWT signing in every test.

### T-PR7-4: `tests/lanzadera/di/test_lanzadera_container.py` (MODIFY)

No direct usage of `auth_bypass` in container tests (the fixture is delivery-only). Verify no tests are broken.

### T-PR7-5: `docs/03-aplicaciones/lanzadera/epic.md` (UPDATE)

Mark W62 as completed in the Quick Navigation table.

### T-PR7-6: Verify

- `grep -r "auth_bypass" tests/ --include="*.py"` → no results.
- `pytest --cov=app --cov-report=term` → 733 total, coverage ≥ 82%.
- `check_test_classification.py --strict` → 0 violations.
- `check_complexity.py --root .` → 0 functions over ceiling.
- `check_mutation_sites.py --root .` → 0 files over ceiling (with BASELINE updates).

### T-PR7-7: Commit + push + PR + merge + cierre del change

Standard. Issue #544 close.

Tras mergear PR-7:
- Cerrar issue #537 (umbrella W62 full auth flow).
- Archivar change: `git mv openspec/changes/w62-auth-flow openspec/changes/archive/w62-auth-flow` (or `openspec archive w62-auth-flow` si el comando existe).
- CHANGELOG: añadir W62 como bloque (1 entrada).

## Phase 6 — Archive (post-merge)

Move `openspec/changes/w62-auth-flow/` to `openspec/changes/archive/`.

## Decisiones operacionales

- Cada PR **parte del branch del PR anterior**, no de main:
  - PR-3 ← origin/a8009bd (main post-PR-2)
  - PR-4 ← origin/feat/540-logout-use-case (PR-3 mergeado)
  - PR-5 ← origin/feat/541-jwt-utility
  - PR-6 ← origin/feat/542-auth-middleware
  - PR-7 ← origin/feat/543-auth-routes

- Conventional commit prefix `feat(lanzadera):` o `fix(lanzadera):`.
- Cada PR usa `size:exception` solo si excede 400 líneas, con `size-exception-reason:` documentado.
- BR-RR: ningún `Co-Authored-By`, ni atribución de IA.
- Branch naming: `feat/<issue>-<kebab-slug>` per `CONTRIBUTING.md`.
- Docker daemon intermittent en CI: rerun con `gh run rerun --failed` o `gh workflow run security.yml --ref <branch>`.

## Lecciones aprendidas (a aplicar en cada PR)

1. **Async fakes**: `await` toda llamada. `_seed_session()` debe `await sessions.create(session)`.
2. **TYPE_CHECKING**: forward refs en `if TYPE_CHECKING:`.
3. **ruff UP037**: no quotes con `from __future__ import annotations`.
4. **mutation_sites BASELINE**: actualizar si el archivo crece.
5. **`_pick` helper**: ya en container.py. Usar para construir adapters.
6. **cast() en properties**: `_user_repo` es `object`; properties usan `cast(X, self._x_repo)`.
7. **asyncio_mode = "auto"**: tests async no necesitan `@pytest.mark.asyncio`.
8. **Docker daemon**: rerun con `gh run rerun --failed` o `gh workflow run security.yml --ref <branch>`.
