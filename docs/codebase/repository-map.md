[← Volver al Codebase Guide](../CODEBASE-GUIDE.md)

# Mapa del repositorio

> Use esta página cuando sepa qué necesita cambiar pero no dónde pertenece.

## Propiedad por paquete

| Ruta | Owns | No ponga aquí |
|---|---|---|
| `app/src/main.py` | Composition root de FastAPI: lifespan, container mirror middleware, healthcheck `/health`, registro de routers `/admin` y `/auth`. | Reglas de negocio o lógica de mutación. |
| `app/src/shared/` | Cross-cutting entre módulos (cache, audit, lockout). | Casos de uso de una app concreta. |
| `app/src/modules/lanzadera/` | MVP de Lanzadera (F01, mother): users, apps, permissions, presence, audit, CLI de bootstrap. | Código de otra app; usar `app/src/modules/<app>/`. |
| `app/src/modules/lanzadera/domain/` | Entidades y value objects: `user.py`, `app.py`, `assignment.py`, `profile.py`, `session.py`, `audit_event.py`, `presence.py`, `global_admin.py`, `reset_token.py`, `legacy_role_map.py`, `errors.py`. Puertos en `app/src/modules/lanzadera/domain/ports/`. | Frameworks, adaptadores, entrega. |
| `app/src/modules/lanzadera/ports/` | Protocolos driving/driven: `cache.py`, `location.py`, `notification_delivery.py`, `bootstrap_admin_source.py`, `jwt_signer.py`, etc. | Implementaciones concretas. |
| `app/src/modules/lanzadera/application/` | Use cases puros: `bootstrap_global_admins.py`, `disable_app.py`, `disable_user.py`, `get_connected_users.py`, `logout.py`, `revoke_assignment.py`, `update_app.py`, `list_effective_apps.py`, `audit_append.py`, `seed_profiles.py`, `issue_reset_token` y `consume_reset_token` (servicios en `app/src/modules/lanzadera/domain/services/`). | Adaptadores o entrega. |
| `app/src/modules/lanzadera/adapters/` | Implementaciones driven: `app/src/modules/lanzadera/adapters/persistence/repositories/<x>_pg.py` (Postgres), `app/src/modules/lanzadera/adapters/cross/secret_manager.py`, `app/src/modules/lanzadera/adapters/crypto/credential_hasher_argon2id.py`, `app/src/modules/lanzadera/adapters/bootstrap/env_admin_source_adapter.py`. | Routers HTTP o CLI. |
| `app/src/modules/lanzadera/di/` | Composition root: `container.py` (LanzaderaContainer), `bootstrap.py` (run_bootstrap), `use_cases.py` (factorías). | UI o escrituras. |
| `app/src/modules/lanzadera/delivery/http/` | Routers FastAPI: `auth_routes.py` (`/auth/login`, `/auth/logout`, `/auth/me`, `/auth/me/apps`), `admin_routes.py` (legacy admin), `admin_routes_users.py`, `admin_routes_users_json.py`, `admin_routes_apps.py`, `admin_routes_misc.py`, `admin_routes_presence.py` (SSE), `admin.py` (records), `auth_middleware.py`, `admin_ports.py`. | Lógica de dominio. |
| `app/src/modules/lanzadera/delivery/http/templates/` | Plantillas Jinja2: `base.html`, `app/src/modules/lanzadera/delivery/http/templates/admin/apps.html`, `app/src/modules/lanzadera/delivery/http/templates/admin/users.html`, `app/src/modules/lanzadera/delivery/http/templates/admin/user_created.html`, `app/src/modules/lanzadera/delivery/http/templates/admin/user_disabled.html`, `app/src/modules/lanzadera/delivery/http/templates/admin/app_activated.html`, `app/src/modules/lanzadera/delivery/http/templates/admin/assignment_created.html`, `app/src/modules/lanzadera/delivery/http/templates/admin/audit.html`. | Lógica de control; queda en los routers. |
| `app/src/modules/lanzadera/delivery/cli/` | Entry points CLI: `platform_user.py` (argparse; subcomandos `set-password`, `grant-global-admin`, `revoke-global-admin`, `list-apps`, `assign-profile`), `platform_user_auth.py`, `platform_user_apps.py`, `platform_user_types.py`, `seed_profiles_cli.py`. | Lógica pura; queda en `application`. |
| `app/src/modules/expedientes/` | Módulo de Expedientes (F04 scaffold): `application` con 30+ use cases (`create_expediente`, `edit_expediente`, `transition_expediente`, `delete_expediente`, `anexos`, `audit`, `authorization`, `autosave_general`, `autosave_related`, `batch_e2e`, `catalogs`, `current_principal`, `e2e_dto`, `e2e_hash`, `e2e_session`, `export_package`, `feedback_guard`, `json_canonical`, `nc`, `ordinal_e2e`, `queries`, `register_anualidad`, `register_hito`, `register_juridica`, `register_responsable`, `register_suministrador`, `riesgos`, `runtime_policies`, `session_state`), `domain` con entidades, `ports`, `migration` (extracción legacy → moderno), `app/src/modules/expedientes/shared/uow.py`. `delivery` está en scaffold (sólo `__init__.py`). | Código de Lanzadera. |
| `app/migrations/versions/` | Migraciones Alembic 0001..0006 (D82, Expand and Contract). | Lógica de aplicación. |
| `app/migrations/env.py` | Runtime Alembic: lee `DATABASE_URL` y aplica `target_metadata`. | Modelos de dominio. |
| `app/pyproject.toml` | Pins exactos de ruff==0.17.2, mypy==1.13.0, argon2-cffi==25.1.0 y del resto del toolchain (Hard Rule 2). | Configuración de runtime. |
| `app/requirements-dev.lock` | Lockfile con hashes de la transitiva (Hard Rule 15). | — |
| `app/requirements.lock` | Lockfile runtime con hashes, export de `app/uv.lock` sin grupos de desarrollo (BP #1456). | — |
| `tests/lanzadera/`, `tests/expedientes/` | Suites TDD por módulo. | Generadores de legacy. |
| `e2e/` | Suite Playwright (F1-F5, 5 fases merged). | Tests unitarios. |
| `scripts/check_*.py` | Quality gates deterministas (ver `docs/gates.md` §2). | Lógica de aplicación. |
| `docs/03-aplicaciones/<app>/` | Research por app: `epic.md`, `walkthrough-G*.json`, `capabilities.md` (el estudio por sistema), `data-model.md`, `forms.md`, `integrations-security.md`, `ERD`. Las fichas de capacidad tienen su ruta canónica en `capabilities/CAP-00X-*.md`, con el índice del sistema en `capabilities/index.md` (#1288). | Código fuente. |
| `data/staging/<app>/` | Binarios `.accdb` R2-pulled (NO se commitean). | Scripts reproducibles; van en `scripts/`. |
| `openspec/changes/<change>/` | Specs SDD: `proposal.md`, `design.md`, `openspec/changes/<change>/specs/<capability>/spec.md`, `tasks.md`. | Implementación. |
| `inputs/` | Material externo versionado ligero (automatizaciones-legacy, etc.). | — |

## Si necesitas X, lee/toca Y

| Si necesita... | Empiece por... | Revise después... |
|---|---|---|
| Cambiar una decisión arquitectónica (D-<n> / DA-<n>) | `docs/decisions/README.md` (una decisión por fichero) | `openspec/changes/<change>/design.md` si está dentro de un change vivo |
| Cambiar un gate de CI | `scripts/<check>.py` (código) | `docs/gates.md` (manual de uso) y `.github/workflows/ci.yml` (qué job lo invoca) |
| Añadir una ruta HTTP de Lanzadera | `app/src/modules/lanzadera/delivery/http/` (router) | `app/src/modules/lanzadera/application/` (use case) y `tests/lanzadera/` (test del flujo) |
| Añadir un subcomando CLI | `app/src/modules/lanzadera/delivery/cli/` (argparse) | `app/src/modules/lanzadera/application/` y `scripts/local-preflight.sh` (smoke) |
| Cambiar la migración de esquema | `app/migrations/versions/<NNNN>_<slug>.py` | `app/migrations/env.py` y `tests/lanzadera/migrations/` |
| Cambiar Argon2id o el reset flow | `app/src/modules/lanzadera/adapters/crypto/credential_hasher_argon2id.py` | `app/src/modules/lanzadera/application/` (`issue_reset_token`, `consume_reset_token`) y `docs/architecture.md` §Auth (D88-D91) |
| Cambiar bootstrap de admins | `app/src/modules/lanzadera/di/bootstrap.py` + `app/src/modules/lanzadera/adapters/bootstrap/env_admin_source_adapter.py` | `scripts/setup-staging.ps1` y `app/src/modules/lanzadera/delivery/cli/platform_user.py` (camino CLI) |
| Añadir un test de Expedientes | `tests/expedientes/application/<use-case>/` | `app/src/modules/expedientes/application/<use-case>/service.py` |
| Cambiar la suite E2E | `e2e/<spec>.spec.ts` | `e2e/helpers/api-client.ts` y `e2e/fixtures/auth-fixture.ts` (globalSetup) |
| Cambiar walkthrough de forms | `docs/03-aplicaciones/<app>/walkthrough-G<N>.json` (data) | `docs/03-aplicaciones/<app>/epic.md` (spec) |
| Cambiar mockups UI | No hay mockups en el repo (ver `docs/codebase/what-is-not-here.md`) | Skill `telefonica-brand-design` y skill `frontend-design` |
| Cambiar la doc de un gates | `docs/gates.md` (manual) | `scripts/<check>.py` (código) y `tests/test_ci_workflow.py` (wiring) |
| Cambiar la suite de skills | `skills/documentation-alan-style/SKILL.md` (estilo) | `skills/README.md` (índice) y `AGENTS.md` (tabla de skills) |

## Regla de colocación

Coloque el código en el owner estable más estrecho:

1. **Una app sólo** → `app/src/modules/<app>/{domain,ports,application,adapters,di,delivery}/`.
2. **Cross-cutting entre apps** → `app/src/shared/` (cache, audit, lockout). NO duplicar por app.
3. **Adaptador Postgres nuevo** → `app/src/modules/<app>/adapters/persistence/repositories/<x>_pg.py` con su tabla en `app/src/modules/<app>/adapters/persistence/repositories/<x>_table.py`.
4. **Caso de uso de Expedientes** → `app/src/modules/expedientes/application/<use-case>/{command,service}.py` (estructura observada en `app/src/modules/expedientes/application/catalogs/`).
5. **Investigación no convertida en código** → `docs/03-aplicaciones/<app>/` o `openspec/changes/<change>/`. No vive en `app/`.
6. **Gate nuevo** → `scripts/check_<name>.py` + entrada en `docs/gates.md` + wiring en `tests/test_ci_workflow.py` + invocación en `.github/workflows/ci.yml::quality` o en `scripts/quality_report.py`.

## Lista de comprobación del contribuidor

- [ ] El cambio tiene un único owner claro (ver tabla de propiedad).
- [ ] NO duplica reglas entre capas: si la lógica vive en `domain`, no se reimplementa en `application` o `adapters`.
- [ ] NO introduce estado generado en `app/src/shared/` que pertenezca a una sola app.
- [ ] Actualiza la doc en el entry point que el usuario va a leer (`docs/gates.md`, `docs/architecture.md`, o el `epic.md` per-app).

## Siguiente paso

Continúe con [Interfaces](interfaces.md) para localizar el router, la plantilla o el comando CLI que su cambio toca.
