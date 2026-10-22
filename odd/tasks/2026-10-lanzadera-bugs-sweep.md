# 2026-10 — Lanzadera Bug Sweep (Closes #773, #822, #774)

## Outcome

Tres PRs desde `main @ HEAD`, cada uno cierra una issue aprobada y no aprobada por la IA:

1. **T1 — Closes #773**: `/auth/logout` invalida el JWT. `AuthMiddleware` consulta el repo de sesiones en cada request autenticada; fail-closed si el repo no responde.
2. **T2 — Closes #822**: ocho `test_DEBUG_*` de auth convertidos a aserción o retirados con nota.
3. **T3 — Closes #774**: `POST /admin/admin/users` acepta body JSON (modelo Pydantic).

Reglas globales: PR ≤400 líneas, `Closes #N` (las tres tienen `status:approved` + `type:bug` — verificado), conventional commits sin atribución de IA, ruff con la config de CI, sin rebase ni force-push, sin `rm -rf`, sin borrar ramas remotas.

## Spec (verbatim de cada issue)

### S1 — issue #773

> `POST /auth/logout` responde 204 (logout "exitoso") pero el JWT sigue siendo válido hasta su `exp` natural. [...] El middleware sigue resolviendo la identidad solo desde el JWT: `auth_middleware.py:92-97` toma `session_id` del claim `sub` y asigna `request.state.user_id = session_id`.

Aceptación (verbatim): "POST `/auth/logout` seguido de `GET /auth/me` con el mismo token devuelve **401**. [...] Si el approach requiere un nuevo decorador / dependencia, se documenta en `docs/architecture.md` §Auth con su D-<n> vigente."

Tests exigidos por el usuario:
- login → /me 200 → logout → /me 401 con el mismo token
- token de sesión inexistente (UUID basura) → /me 401
- token de sesión expirada → /me 401
- repo de sesiones no responde → /me 401 (fail-closed; NUNCA acepta)

Manejo del bloqueante PR #934: la fix vive entera en `auth_middleware.py` (no toca `main.py`). El middleware resuelve `request.app.state.container.sessions` en `dispatch`, igual que PR #934 ya resuelve `app.state.container.jwt_signer` en su `_LazyAuthSigner`. Si más adelante aparece la necesidad de tocar `main.py`, pausa para merge de #934.

### S2 — issue #822

> `tests/lanzadera/delivery/test_auth_routes_integration.py` contiene ocho pruebas `test_DEBUG_*` que solo imprimen rutas, fuentes o fechas de fichero: **ninguna tiene un solo `assert`**.

Aceptación (verbatim): "Ninguna función `test_*` del fichero termina sin al menos una aserción: o imprime y afirma, o se retira. [...] Lo que hoy se inspecciona a ojo y merezca cobertura (las rutas registradas de auth, la correspondencia fuente/bytecode) tiene su aserción, o queda escrito dónde se cubre."

Decisión caso-por-caso en este PR (se documenta en el cuerpo):

| Test | Decisión | Justificación |
|---|---|---|
| `test_DEBUG_print_routes` | RETIRAR | `_fakes` + `auth_client` ya prueban rutas por sus tests funcionales. |
| `test_DEBUG_source_contains_my_apps` | RETIRAR | `test_my_apps_returns_*` cubren el comportamiento. |
| `test_DEBUG_call_register` | RETIRAR | Sustituido por los tests de cada ruta (`/auth/me`, `/auth/me/apps`, `/auth/me/apps/{app_id}/capabilities`). |
| `test_DEBUG_patch_then_register` | RETIRAR | `test_my_apps_*` + coverage del path cubren el re-register. |
| `test_DEBUG_reload_and_register` | RETIRAR | El reload es operación de dev-only; la cobertura de las rutas ya basta. |
| `test_DEBUG_source_vs_bytecode` | RETIRAR | Diagnóstico del fix PR, ya consumido por la suite. |
| `test_DEBUG_raw_source` | RETIRAR | Sustituido por `test_my_apps_returns_apps_with_profiles_and_capabilities` y similares. |
| `test_DEBUG_file_mtime` | RETIRAR | El `.pyc` vs `.py` no es contrato. Si fuera problema, está cubierto por `scripts/quality_report.py`. |

Si en el review Alan pide mantener alguno como aserción, se convierte in-place: `assert isinstance(routes, list)` o `assert "my_apps" in src`.

### S3 — issue #774

> El handler está declarado así: `*, email: str, name: str, national_id: str` — sin `Body()` o modelo Pydántico convierte los parámetros en **query params obligatorios**. Resultado: el endpoint rechaza el body JSON con 422.

Aceptación (verbatim): "El handler queda declarado con un modelo Pydántico `CreateUserRequest` (en `app/src/modules/lanzadera/delivery/http/schemas.py` o equivalente). [...] El router legacy (`admin_routes_users.py`) mantiene la API con query string."

## Branch & worktree

Tres ramas, una por PR (regla `chained-pr` no aplica — son cambios independientes sobre `main`):

| PR | Issue | Branch desde `main` | Worktree path |
|---|---|---|---|
| T1 | #773 | `fix/773-auth-logout-invalidates-jwt` | `<repo>/../access2web-blueprint-worktrees/773-auth-logout-jwt` |
| T2 | #822 | `chore/822-test-debug-no-asserts` | `<repo>/../access2web-blueprint-worktrees/822-test-debug-asserts` |
| T3 | #774 | `fix/774-admin-users-json-body` | `<repo>/../access2web-blueprint-worktrees/774-admin-users-body` |

Sincronización intra-PR: `git fetch origin && git merge origin/main` dentro de la rama (regla 1 de `governed-delivery`).

## Patrones

### T1 — AuthMiddleware consulta el repo de sesiones

```python
# auth_middleware.py — dentro de dispatch, tras extraer session_id del JWT
container = getattr(request.app.state, "container", None)
session_repo = getattr(container, "sessions", None) if container is not None else None

if session_repo is not None:
    try:
        session = await session_repo.get_by_id(session_id)
    except Exception:  # noqa: BLE001 — fail-closed explícito
        payload = None   # trato equivalente a "no autenticado"
    else:
        if session is None:
            payload = None
        elif session.expires_at <= datetime.fromtimestamp(self._now(), UTC):
            payload = None
```

Reglas:
- Si `container` o `session_repo` no están disponibles (pre-lifespan, tests sin container) → comportamiento actual: dejar `state.user_id = None`.
- Si `get_by_id` lanza → fail-closed (`payload = None`, state vacío → `require_*` devuelve 401/403).
- Si la sesión expiró → state vacío (consistente con expiración del JWT).
- No se introduce un decorador nuevo (Pydantic invalidaría `Request`, sería scope creep); basta con un método interno.

### T1 — test RED (integration)

Archivo nuevo: `tests/lanzadera/delivery/test_auth_logout_invalidates_token.py`, marcado `pytest.mark.integration`.

Casos (cada uno: RED → commit → GREEN → commit):

1. `test_logout_then_me_returns_401` — login → /me 200 → logout → /me 401 (mismo token).
2. `test_me_with_unknown_session_id_returns_401` — JWT con `sub=<uuid aleatorio no presente en sessions>` → /me 401.
3. `test_me_with_expired_session_returns_401` — JWT firmado con `exp` futura pero fila en DB con `expires_at` ya pasado (escenario distinto del JWT expirado, captura la rama de expires_at check del repo).
4. `test_me_returns_401_when_session_repo_unreachable` — monkey-patch `container.sessions.get_by_id` para que lance `RuntimeError`; /me 401. Estado: el container queda igual para no contaminar a otros tests (fixture `autouse` cleanup).

Setup: `APAP_INTEGRATION_ENABLED=1` + `APAP_INTEGRATION_URL=postgresql+asyncpg://...`. Cada test siembra su usuario y su sesión (cleanup en teardown; cascada por FK).

### T2 — retirar ocho `test_DEBUG_*`

Una sola edit en `tests/lanzadera/delivery/test_auth_routes_integration.py`: borrar las ocho funciones (líneas 668-843 aprox) y el comentario introductorio. Cuerpo del PR incluye la tabla `Decisión caso-por-caso` arriba.

Verificación: `python scripts/check_test_classification.py --strict` y `rg -n "def test_" tests/lanzadera/delivery/test_auth_routes_integration.py` muestra cero `test_DEBUG_*`.

### T3 — Pydantic body model para `POST /admin/users`

```python
# app/src/modules/lanzadera/delivery/http/schemas.py (nuevo)
class CreateUserRequest(BaseModel):
    email: EmailStr
    name: str = Field(min_length=1, max_length=200)
    national_id: str = Field(min_length=1, max_length=32)

# admin_routes_users_json.py — handler
@router.post("", status_code=status.HTTP_201_CREATED)
async def create_user(
    request: Request,
    body: CreateUserRequest,   # FastAPI lo lee del body
) -> dict[str, object]:
    ...
    user = await container.create_user(
        email=body.email.strip().lower(),
        name=body.name.strip(),
        national_id=body.national_id.strip(),
        actor_id=actor,
    )
```

Test nuevo en `tests/lanzadera/delivery/test_admin_users_json_create.py` con `@pytest.mark.integration`:
- `test_create_user_with_json_body_returns_201` — POST con JSON real + bearer válido.
- `test_create_user_with_missing_field_returns_422` — body sin `name` → 422.
- `test_create_user_with_query_only_still_accepted_or_422` — verifica el contrato del cambio.

## Hard rules (comunes a las tres PRs)

- Conventional commits sin atribución de IA. Mensajes:
  - T1: `fix(lanzadera): invalidate JWT on logout (AuthMiddleware consults session repo) (Closes #773)`
  - T2: `chore(tests): remove eight test_DEBUG_* auth tests that print without asserting (Closes #822)`
  - T3: `fix(lanzadera): accept JSON body in POST /admin/users (Pydantic CreateUserRequest) (Closes #774)`
- `from __future__ import annotations` consistente con los archivos editados.
- Sin nuevos contratos públicos, sin tocar otras apps/módulos.
- Ruff con la config de CI:
  ```bash
  ruff check --config app/pyproject.toml <paths>
  ruff format --check --config app/pyproject.toml <paths>
  ```
- PR size: ≤400 líneas (`python3 scripts/check_pr_size.py --base-ref main`).
- Sin rebase ni force-push, sin `rm -rf`.

## Verificación local pre-PR (común)

```bash
cd <worktree-dir>
# Sincronizar main
git fetch origin
git merge origin/main   # ff-only en el primer commit; merge en updates

# Calidad
ruff check --config app/pyproject.toml <paths-tocados>
ruff format --check --config app/pyproject.toml <paths-tocados>
mypy --explicit-package-bases <paths-tocados>
python scripts/check_branch_name.py <nombre-rama>
python scripts/check_pr_size.py --base-ref main

# Suite relevante
pytest -c app/pyproject.toml --rootdir=app -m "not integration and not e2e and not e2e_smoke" tests/lanzadera/delivery/
# Para T1 además (con Postgres local):
APAP_INTEGRATION_ENABLED=1 APAP_INTEGRATION_URL=postgresql+asyncpg://lanzadera:lanzadera@127.0.0.1:5432/lanzadera \
  pytest -c app/pyproject.toml --rootdir=app -m integration tests/lanzadera/delivery/test_auth_logout_invalidates_token.py
```

## Estado

- [ ] T1 worktree creado desde `main`
- [ ] T1 Postgres local corriendo (docker run postgres:16-alpine) + schema aplicado
- [ ] T1 tests RED escritos (4 casos) — `pytest -m integration tests/lanzadera/delivery/test_auth_logout_invalidates_token.py` falla
- [ ] T1 AuthMiddleware extendido — tests GREEN
- [ ] T1 ruff + mypy + check_pr_size verdes
- [ ] T1 commit + push + PR abierto con "Closes #773"
- [ ] T2 worktree creado desde `main` actualizado post-T1-merge
- [ ] T2 ocho `test_DEBUG_*` retirados (1 commit)
- [ ] T2 quality_report + check_test_classification verdes
- [ ] T2 commit + push + PR abierto con "Closes #822"
- [ ] T3 worktree creado desde `main` actualizado post-T2-merge
- [ ] T3 schemas.py + handler modificado + tests integración (2-3 casos)
- [ ] T3 ruff + mypy + check_pr_size verdes
- [ ] T3 commit + push + PR abierto con "Closes #774"
