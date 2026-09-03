# Design: W62 auth flow (PR-3..PR-7)

> **Change**: `w62-auth-flow` · **Project**: `access2web-blueprint`

## Architecture overview

```
                            ┌─────────────────┐
HTTP request                │ FastAPI app      │
  Authorization: Bearer ───▶│ AuthMiddleware    │  PR-5
  (header)                  │ (Starlette ASGI) │
                            └────────┬────────┘
                                     │ request.state.user_id = <session_id>
                                     ▼
┌──────────────────────────────────────────────────────────────────────┐
│ Auth routes (PR-6)         │ Admin routes (PR-7)                    │
│  POST /auth/login          │  POST /admin/users                    │
│  POST /auth/logout         │  PATCH /admin/apps/{id}               │
│  GET /auth/me              │  DELETE /admin/apps/{id}              │
│                            │                                        │
│  → container.login()       │  → require_global_admin(read state)   │
│  → container.logout()      │  → _actor_id(request) ← state.user_id│
│                            │  → container.create_user(actor_id)    │
└────────┬───────────────────┴──────────┬─────────────────────────────┘
         │                              │
         ▼                              ▼
┌──────────────────────────────────────────────────────────────────────┐
│ LanzaderaContainer (DI)                                                  │
│  use_cases["login"]   → application/login.py          (PR-2)        │
│  use_cases["logout"]  → application/logout.py         (PR-3)        │
│  use_cases["create_user"], ["disable_user"], ["update_app"], etc.     │
│  jwt_signer            → adapters/crypto/jwt.py        (PR-4)        │
└────────┬───────────────────┬─────────────────────────────────────────┘
         │                   │
         ▼                   ▼
┌─────────────────┐  ┌─────────────────────────────────┐
│ SessionRepo     │  │ UserRepo / AppRepo / ...          │
│ (PR-1)          │  │ (Postgres adapters / fakes)       │
│ create/get/revoke│ │                                 │
└─────────────────┘  └─────────────────────────────────┘
```

## Decisiones arquitectónicas

### AD-W62-1: Sin librerías JWT externas

**Decisión**: implementar JWT con stdlib (`hmac`, `hashlib`, `base64`, `secrets`).

**Rationale**: minimizar superficie de seguridad. Las libs JWT (PyJWT, authlib) son confiables pero añaden 2-4 deps transitivas. HS256 con HMAC-SHA256 stdlib es ~80 líneas y es estándar (RFC 7515).

**Alternativas consideradas**:
- PyJWT: estándar, pero añade dep + 4 transitivas. Aceptable si en el futuro hace falta RS256.
- authlib: overkill para HS256 simple.

**Revisable si**: el auth flow crece a RS256, JWKS, o rotación de keys.

### AD-W62-2: Middleware silent-fail

**Decisión**: el middleware no raise cuando el token falta o es inválido. Deja `request.state.user_id = None` y deja pasar el request. Las rutas destructivas usan `require_global_admin()` que sí retorna 401/403.

**Rationale**: el middleware no es el gate. El gate vive en cada ruta (depende de la policy). Las rutas públicas (`/health`, `/auth/login`) no necesitan auth, así que el middleware debe dejarlas pasar sin auth.

**Alternativas**:
- Middleware raise → 401 antes de cualquier ruta. Inconveniente: bloquea rutas públicas.
- Middleware redirect → 302 a /login. Inconveniente: requiere HTML.

### AD-W62-3: Auth tests con test-only middleware

**Decisión**: los admin integration tests usan un middleware de test (`X-Test-Session-Id → request.state.user_id`) en lugar del `AuthMiddleware` real.

**Rationale**: las admin routes ejercitan `require_global_admin`, que es el contrato que queremos verificar. El JWT signing es responsabilidad del middleware (PR-5); los admin routes no deberían depender de él para sus tests. El test-only middleware aísla el contrato de cada capa.

**Alternativa**: firmar JWT real en cada admin integration test. Inconveniente: añade boilerplate sin probar nada nuevo.

### AD-W62-4: Chain PR-by-PR

**Decisión**: 5 PRs encadenados (PR-3..PR-7), cada uno con su propio diff < 400 líneas + size:exception documentado si excede.

**Rationale**: PRs más pequeños = review más rápido. La cadena es tight (cada PR depende del anterior), pero la chain strategy es la convención del repo (CONTRIBUTING.md §Tamaño de los PRs).

**Tamaños esperados**:
| PR | Líneas estimadas | size:exception probable |
|---|---|---|
| PR-3 logout | ~300 | No |
| PR-4 JWT | ~250 | No |
| PR-5 middleware | ~250 | No |
| PR-6 auth routes | ~600 | Sí (auth_routes + admin.py + container + main.py) |
| PR-7 stub replacement | ~200 | No |

### AD-W62-5: PR-6 también actualiza main.py

**Decisión**: PR-6 incluye el wiring en `app/src/main.py` (entry point de la app FastAPI). Esto es parte del chain porque main.py no se prueba hasta que el middleware + las rutas existen.

**Rationale**: chain de middleware → container → routes → main wiring es coherente. Splittar main.py en otro PR introduce un commit "no-op" (main.py importa cosas que no existen todavía).

**Decisión alternativa (descartada)**: PR-5.5 "wire main.py" separado. Inconveniente: review workload overhead.

## Cross-cutting concerns

### D-7 (abierta, no se cierra en este change)

D-7 — ubicación de tests de futuras apps — sigue trigger-cerrada (#534). No se aborda en W62.

### G1/G2/G3 (gaps de testing-strategy)

G1 (Postgres e2e en CI), G2 (concurrencia), G3 (atomicidad audit) siguen abiertos. No se atacan en W62.

### Mutation sites BASELINE

PR-4 (`Hs256JwtSigner`) y PR-5 (`auth_middleware.py`) crearán archivos nuevos. Si exceden 100 sites, agregar al BASELINE con `target_date="2027-02-13"`. Patrón establecido en PR-2.

### Complexity

Las `auth_routes.py` (PR-6) puede tener una función con CC > 10 si se acumulan las branches de error. Refactor pattern: extraer `_format_login_error(exc)` helper, similar al `_pick` helper de PR-2.

## Sequencing

```text
PR-3 (logout)              ──┐
PR-4 (JWT)                 ──┤
PR-5 (middleware)          ──┤
PR-6 (auth routes + main)  ──┤── todos mergeados secuencialmente
PR-7 (stub replacement)   ──┘
```

Cada PR parte del branch del anterior (no de main). `git worktree add` con `-b feat/<n>-... origin/<prev-branch>`.

## Verification matrix

| Layer | Tool | Threshold |
|---|---|---|
| Domain ports (Category 3) | pytest | contract tests |
| Use cases (Category 2) | pytest | FakeFixtures |
| HTTP delivery (Category 4) | pytest + httpx.AsyncClient | TestClient / ASGITransport |
| Type checks | mypy | strict, --explicit-package-bases |
| Format | ruff format | line-length 100 |
| Lint | ruff check | select=E,F,W,I,UP,B |
| Complexity | scripts/check_complexity.py | CC ≤ 10 (per function) |
| Mutation sites | scripts/check_mutation_sites.py | ≤ 100 per file (BASELINE ratchet) |
| Taxonomy | scripts/check_test_classification.py | 5 categorías |
| DRY | scripts/check_dry.py | BASELINE ratchet |
| Layer purity | scripts/check_layers.py | hexagonal purity |

## Cross-references

- `docs/03-aplicaciones/lanzadera/epic.md` — epic madre.
- `openspec/changes/lanzadera-mvp/specs/auth-core/spec.md` — credential storage, lockout (overlaps with W62).
- `openspec/changes/lanzadera-mvp/specs/auth-bootstrap/spec.md` — bootstrap CLI (tangential).
- `tests/lanzadera/_fakes.py` — fakes shared across all tests.
- `skills/lanzadera-testing-strategy/SKILL.md` — taxonomía de las 5 categorías.
- `scripts/check_test_classification.py` — gate de la taxonomía (PR-2 del testing-strategy epic).
