[← Volver al Codebase Guide](../CODEBASE-GUIDE.md)

# Interfaces

> Esta página rastrea `main`, que puede contener cambios sin liberar. Para la última versión liberada, véase la etiqueta correspondiente.

Este repo expone **HTTP (FastAPI + Jinja2)**, **CLI de plataforma** y **scripts operativos**. No expone una API HTTP local en un puerto separado, ni un dashboard SPA, ni un servidor gRPC.

## Mapa de interfaces

| Interface | Estado en este repo | Ficheros primarios | Detalle |
|---|---|---|---|
| HTTP API (JSON) | Implementada (MVP de Lanzadera) | `app/src/main.py`, `app/src/modules/lanzadera/delivery/http/admin_routes_apps.py`, `admin_routes_users_json.py`, `admin_routes_presence.py`, `auth_routes.py` | [Rutas HTTP](#rutas-http) abajo |
| HTTP HTML (HTMX) | Implementada (admin + auth) | `app/src/modules/lanzadera/delivery/http/admin_routes_*.py`, `app/src/modules/lanzadera/delivery/http/templates/admin/*.html` | [Plantillas](#plantillas) abajo |
| CLI de plataforma | Implementada (5 subcomandos) | `app/src/modules/lanzadera/delivery/cli/platform_user.py`, `platform_user_auth.py`, `platform_user_apps.py`, `platform_user_types.py`, `seed_profiles_cli.py` | [CLI de plataforma](#cli-de-plataforma) abajo |
| Scripts operativos | Implementados | `scripts/check_*.py`, `scripts/local-preflight.sh`, `scripts/release-preflight.sh`, `scripts/quality_report.py`, `scripts/mutation_shards.py`, `scripts/setup-staging.ps1`, `scripts/sync-to-r2.ps1`, `scripts/migrate_from_access.py` | [Scripts](#scripts) abajo |
| API HTTP local en puerto separado | No presente | Ningún paquete de dashboard o servidor separado | Use la propia FastAPI app (puerto único) |
| Dashboard SPA | No presente | No existe un bundle JS de UI | Use las plantillas Jinja2 + HTMX + Alpine.js |
| gRPC / WebSocket persistente | No presente | Sólo SSE puntual en `/admin/presence/stream` | Si necesita más, abra una epic en `openspec/changes/<change>/` |

## Rutas HTTP

El composition root en `app/src/main.py` monta los routers de Lanzadera con los prefijos `/auth` y `/admin`. Los routers `admin_routes_*` se registran dentro del lifespan para evitar la circular import con `app.state.container` (W62, #543; e2e-smoke, #876).

### Auth (`app/src/modules/lanzadera/delivery/http/auth_routes.py`)

| Método | Ruta | Fichero | Notas |
|---|---|---|---|
| `POST` | `/auth/login` | `auth_routes.py` | Body JSON; emite JWT. |
| `POST` | `/auth/logout` | `auth_routes.py` | Invalida el JWT actual; `#773` fija el contrato. |
| `GET` | `/auth/me` | `auth_routes.py` | Devuelve el usuario actual. |
| `GET` | `/auth/me/apps` | `auth_routes.py` | Lista de aplicativos efectivos. |
| `POST` | `/auth/reset/request` | `auth_routes.py` | Solicita reset (issue #885, dominio). |
| `POST` | `/auth/reset/consume` | `auth_routes.py` | Consume token; ver `app/src/modules/lanzadera/domain/services/issue_reset_token.py` y `app/src/modules/lanzadera/domain/services/consume_reset_token.py`. |

### Admin JSON (`admin_routes_users_json.py` + `admin_routes_apps.py`)

| Método | Ruta | Fichero | Notas |
|---|---|---|---|
| `GET` | `/admin/users` | `admin_routes_users_json.py` | Lista paginada de usuarios. |
| `POST` | `/admin/users` | `admin_routes_users_json.py` | Crea usuario (201). |
| `GET` | `/admin/users/{user_id}` | `admin_routes_users_json.py` | Detalle de usuario. |
| `PATCH` | `/admin/users/{user_id}` | `admin_routes_users_json.py` | Actualización parcial. |
| `DELETE` | `/admin/users/{user_id}/assignments/{app_id}` | `admin_routes_users_json.py` | Revoca asignación (204). |
| `POST` | `/admin/apps` | `admin_routes_apps.py` | Crea aplicativo (201). |
| `GET` | `/admin/apps/{app_id}` | `admin_routes_apps.py` | Detalle de aplicativo. |
| `PATCH` | `/admin/apps/{app_id}` | `admin_routes_apps.py` | Actualización parcial. |
| `DELETE` | `/admin/apps/{app_id}` | `admin_routes_apps.py` | Baja (204). |
| `GET` | `/admin/apps/{app_id}/records` | `admin.py` | Registros legacy. |

### Admin HTML (`admin_routes_*.py` + `app/src/modules/lanzadera/delivery/http/templates/admin/`)

| Método | Ruta | Fichero de router | Plantilla | Notas |
|---|---|---|---|---|
| `GET` | `/admin/apps` | `admin_routes_misc.py` | `app/src/modules/lanzadera/delivery/http/templates/admin/apps.html` | Listado HTMX. |
| `POST` | `/admin/apps/{app_id}/activate` | `admin_routes_misc.py` | `app/src/modules/lanzadera/delivery/http/templates/admin/app_activated.html` | Activación. |
| `POST` | `/admin/assignments` | `admin_routes_misc.py` | `app/src/modules/lanzadera/delivery/http/templates/admin/assignment_created.html` | Crea asignación (201). |
| `GET` | `/admin/audit` | `admin_routes_misc.py` | `app/src/modules/lanzadera/delivery/http/templates/admin/audit.html` | Eventos de auditoría. |
| `GET` | `/admin/users` | `admin_routes_users.py` | `app/src/modules/lanzadera/delivery/http/templates/admin/users.html` | Listado. |
| `POST` | `/admin/users` | `admin_routes_users.py` | `app/src/modules/lanzadera/delivery/http/templates/admin/user_created.html` | Alta (201). |
| `PATCH` | `/admin/users/{user_id}/disable` | `admin_routes_users.py` | `app/src/modules/lanzadera/delivery/http/templates/admin/user_disabled.html` | Desactivación. |
| `POST` | `/admin/presence/heartbeat` | `admin_routes_presence.py` | — | Latido de presencia (#771). |
| `GET` | `/admin/presence` | `admin_routes_presence.py` | — | Snapshot. |
| `GET` | `/admin/presence/stream` | `admin_routes_presence.py` | — | SSE puntual; `presence_stream` resuelve el container vía `request.state.container` (W60, #522). |

### Healthcheck

| Método | Ruta | Fichero | Notas |
|---|---|---|---|
| `GET` | `/health` | `app/src/main.py` | `{"status":"ok","phase":"1","service":"lanzadera"}`; usado por `e2e` y por el Docker preflight (`.github/workflows/release.yml::e2e`). |

### Montaje

El orden de `app.include_router` en `app/src/main.py` (auth → presence → apps → users_json → admin) es determinista: FastAPI dispatch por match exacto, así que los slices no colisionan en rutas compartidas. Mantenga ese orden en PRs que añadan routers nuevos.

## Plantillas

Jinja2 vive bajo `app/src/modules/lanzadera/delivery/http/templates/`. El composition root fija el directorio con `Jinja2Templates(directory=str(_TEMPLATES_DIR))` y los routers lo reciben por inyección.

| Plantilla | Sirve a | Notas |
|---|---|---|
| `base.html` | Layout común | Heredada por todas las `app/src/modules/lanzadera/delivery/http/templates/admin/*.html`. |
| `app/src/modules/lanzadera/delivery/http/templates/admin/apps.html` | GET `/admin/apps` | Listado HTMX. |
| `app/src/modules/lanzadera/delivery/http/templates/admin/users.html` | GET `/admin/users` | Listado HTMX. |
| `app/src/modules/lanzadera/delivery/http/templates/admin/user_created.html` | POST `/admin/users` | Confirmación parcial HTMX. |
| `app/src/modules/lanzadera/delivery/http/templates/admin/user_disabled.html` | PATCH `/admin/users/{user_id}/disable` | Confirmación parcial. |
| `app/src/modules/lanzadera/delivery/http/templates/admin/app_activated.html` | POST `/admin/apps/{app_id}/activate` | Confirmación parcial. |
| `app/src/modules/lanzadera/delivery/http/templates/admin/assignment_created.html` | POST `/admin/assignments` | Confirmación parcial. |
| `app/src/modules/lanzadera/delivery/http/templates/admin/audit.html` | GET `/admin/audit` | Tabla de eventos. |

Si añade una plantilla nueva, declárela en la tabla de arriba y mantenga el tono castellano peninsular formal con usted (ver `skills/documentation-alan-style/SKILL.md`).

## CLI de plataforma

`app/src/modules/lanzadera/delivery/cli/platform_user.py` define `argparse` con cinco subcomandos. El CLI es el camino de bootstrap (`set-password` es el único camino que crea la fila inicial del admin global; D26 obliga a confirmación interactiva en cualquier subcomando destructivo).

| Subcomando | Fichero | Args | Notas |
|---|---|---|---|
| `set-password` | `platform_user_auth.py` | `email` | Bootstrap inicial. |
| `grant-global-admin` | `platform_user_auth.py` | `email` | Concesión. |
| `revoke-global-admin` | `platform_user_auth.py` | `email` | Revocación. |
| `list-apps` | `platform_user_apps.py` | — | Listado. |
| `assign-profile` | `platform_user_apps.py` | `email`, `app_id`, `profile_code` | Asignación perfil-app. |

`seed_profiles_cli.py` carga los perfiles del seed inicial (ver `app/migrations/versions/0003_seed_profiles.py` para el seed Alembic equivalente).

Para invocar: `python -m app.src.modules.lanzadera.delivery.cli.platform_user <subcomando> ...` desde la raíz del repo.

## Scripts

`scripts/` agrupa quality gates y tool scripts. Los quality gates están documentados en `docs/gates.md`; aquí solo se lista el contrato de superficie.

| Script | Tipo | Quién lo invoca | Notas |
|---|---|---|---|
| `.importlinter` | Gate | `ci.yml::docs-architecture` | DA-1; derivado del bloque `architecture`, rechaza imports que violen las capas. |
| `scripts/check_legacy_hashes.py` | Gate | `ci.yml::quality` | DA-13; prohíbe `legacy_hash`, `sha256`, `migrate_password`. |
| `scripts/check_mutation_sites.py` | Gate | `ci.yml::quality` | DA-13; pin test AST. |
| `scripts/check_complexity.py` | Gate | `ci.yml::quality` | Complejidad ciclomática. |
| `scripts/check_dry.py` | Gate | `ci.yml::quality` | DRY. |
| `scripts/check_deadcode.py` | Gate | `ci.yml::quality` | Código muerto. |
| `scripts/check_pr_size.py` | Gate | `ci.yml::pr-gate` | ≤400 líneas. |
| `scripts/check_workflows.py` | Gate | `ci.yml::quality` | Valida `.github/workflows/*.yml`. |
| `.github/required-jobs-policy.json` | Política del gate del canon | `ci.yml::required` | Declara la matriz `required` y los saltos aceptados por evento. |
| `scripts/check_decision_guards.py` | Gate | `ci.yml::quality` | Decisiones D-<n> con guardia. |
| `scripts/check_walkthrough_schema.py` | Gate | `ci.yml::pr-gate` | `docs/03-aplicaciones/<app>/walkthrough-G*.json`. |
| `.agents/skills/ci-pattern/assets/issue-spec/check_issue_spec.py` | Gate del canon | `ci.yml::pr-gate` | Cuerpo de issue conforme a `docs/08-decisiones-y-preguntas-abiertas.md`. |
| `scripts/check_suites_populated.py` | Gate | `ci.yml::quality` | Suites obligatorias no vacías. |
| `scripts/check_test_classification.py` | Gate | `ci.yml::quality` | Marca tests según marcador. |
| `scripts/check_legacy_retirement.py` | Gate | `ci.yml::quality` | Retirada de paths legacy. |
| `scripts/check_mutation.py` | Gate | `release.yml::mutation` | Ratchet shrink-only sobre cosmic-ray. |
| `scripts/local-preflight.sh` | Tool | Local (desarrollador) | Smoke antes de push. |
| `scripts/release-preflight.sh` | Tool | `release.yml::preflight` | Verifica identidad del tag. |
| `scripts/require-ci-success.sh` | Tool | `release.yml::preflight` | Exige CI verde en el SHA. |
| `scripts/quality_report.py` | Tool | `ci.yml::quality` (step final) | Renderiza el quality report a markdown. |
| `scripts/mutation_shards.py` | Tool | `release.yml::mutation` | Reparte mutantes en 10 tramos deterministas. |
| `scripts/setup-staging.ps1` | Tool | Local (desarrollador) | Pull de `data/staging/<app>/` desde R2. |
| `scripts/sync-to-r2.ps1` | Tool | Local (desarrollador) | Sube binarios a R2. |
| `scripts/migrate_from_access.py` | Tool | Local (desarrollador) | Migración legacy → hexagonal. |
| `scripts/migrate_walkthrough_schema.py` | Tool | Local (desarrollador) | Migración de `walkthrough-*.json`. |
| `scripts/generate_lanzadera_fixtures.py` | Tool | Local (desarrollador) | Genera fixtures de tests. |

`scripts/install-skills.sh` y `scripts/install-skills.ps1` instalan la copia internalizada de skills del repo (ver `skills/README.md`).

## Cómo añadir una ruta, plantilla o subcomando

1. Identifique la capa correcta en [Mapa del repositorio](repository-map.md).
2. Si es HTTP JSON → router en `app/src/modules/lanzadera/delivery/http/`, use case en `application`, test en `tests/lanzadera/delivery/` o `tests/lanzadera/application/`.
3. Si es HTTP HTML → router + plantilla nueva en `app/src/modules/lanzadera/delivery/http/templates/admin/`; actualice esta tabla.
4. Si es CLI → `app/src/modules/lanzadera/delivery/cli/`; añada `sub.add_parser(...)` en `build_parser()` y registre la rama en `scripts/check_test_classification.py`.
5. Si es script → `scripts/<name>.py` o `.sh`; añada el wiring en `.github/workflows/*.yml` y la entrada en `docs/gates.md` si es un gate.

## Fronteras explícitas

- **No exponga endpoints nuevos sin actualizar esta tabla.** Si lo hace, una PR de revisión debe exigir la entrada correspondiente.
- **No duplique rutas entre routers.** El orden en `app/src/main.py` es determinista; cambiarlo requiere justificación en la PR y un test en `tests/lanzadera/delivery/`.
- **No exponga un puerto separado.** Toda la superficie HTTP va por la app FastAPI única (`/auth`, `/admin`, `/health`). Un puerto nuevo abre un riesgo de red distinto al que el repo ha modelado.
- **No invoque los CLI desde los routers.** Los CLI son paths de bootstrap, no API. Si una vista necesita la misma operación, expóngala como endpoint (router + use case + test).

## Si necesitas X, lee/toca Y

| Si necesitas... | Empieza por... | Luego revisa... |
|---|---|---|
| Añadir o cambiar una ruta HTTP de administración | `app/src/modules/lanzadera/delivery/http/admin_routes.py` | `app/src/modules/lanzadera/delivery/http/templates/admin/` y `docs/codebase/testing-and-ci.md` |
| Cambiar la autenticación HTTP | `app/src/modules/lanzadera/delivery/http/auth_routes.py` | `app/src/modules/lanzadera/delivery/http/auth_middleware.py` |
| Añadir un subcomando de CLI | `app/src/modules/lanzadera/delivery/cli/platform_user.py` | `app/src/modules/lanzadera/delivery/cli/` |
| Añadir un gate de CI | `scripts/check_complexity.py` como modelo | `docs/gates.md` y `.github/workflows/ci.yml` |

## Regla de colocación

Las rutas y plantillas HTTP van en el paquete `delivery` del módulo dueño; la lógica de negocio nunca va en una ruta, sino en `application`.

## Lista de comprobación del contribuidor

- [ ] La ruta, plantilla, subcomando o script tiene un único owner claro.
- [ ] Si añadió una ruta HTTP, su test vive en `tests/lanzadera/delivery/` o `tests/lanzadera/application/` con `@pytest.mark.integration` si toca Postgres.
- [ ] Si añadió una plantilla, mantiene el layout de `base.html` y el tono castellano peninsular formal.
- [ ] Si añadió un script, lo añadió a `docs/gates.md` (si es gate) o a `.github/workflows/*.yml` (si lo invoca CI) y a esta tabla.

## Siguiente paso

Continúe con [Testing y CI](testing-and-ci.md) para saber qué corre en PR, nightly y release.