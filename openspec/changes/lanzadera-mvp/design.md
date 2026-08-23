<!-- Change: lanzadera-mvp · Project: access2web-blueprint · Date: 2026-08-08 -->

# Design: Lanzadera MVP — primer slice de la plataforma hexagonal

> **Sentence que organiza**: la hexagonalidad no se discute, se verifica. Cada decisión arquitectónica aprobada se traduce en una forma física, un puerto y un gate que falla cuando el código la viola. El diseño del módulo `lanzadera` se ata a D5–D91 y a QC-1 a QC-11 sin invenciones nuevas.

## Resumen ejecutivo

Este diseño describe cómo se construye el primer slice del monolito modular `platform/` sobre PostgreSQL, partiendo del módulo `lanzadera` y de los nueve sub-specs ya producidos por `sdd-spec` (`users`, `apps`, `profiles`, `assignments`, `auth-core`, `auth-reset`, `auth-bootstrap`, `global_admins`, `audit`). Las decisiones arquitectónicas heredadas son D5 (plataforma modular permission-aware), D8 (hexagonal global), D36+D37 obsoletos, D66-D68 (stack cerrado y monolito modular), D70-D72 (caché selectiva), D77 (Docker desde día uno), D82 (Expand and Contract), D85 (catálogo de 20 aplicaciones), D88 (Argon2id vía `argon2-cffi==25.1.0`), D89 (sin columna `legacy_hash`, `password_hash = NULL` para los 156 usuarios), D90 (reset flow con tokens one-time de 24 h) y D91 (CLI exclusivo para el primer admin global). Los doce quality gates QC-1 a QC-11 aparecen wired en `.github/workflows/ci.yml`, pinned por `tests/test_ci_workflow.py` y commiteados antes del primer `git commit` de código de aplicación, conforme al plan de `docs/calidad-de-codigo-y-ci.md`. La entrega se reparte en cuatro fases de cero a dos semanas, con la fase de aplicación (Alembic 0001-0006, adapters de repos, delivery HTMX, CLI admin) ejecutada en strict TDD.

## Estructura física del módulo

La forma hexagonal del módulo sigue el monolito modular aprobado por D68 y los límites por paquete del gate `scripts/check_layers.py`. El root package para el gate es `app.src.modules`; las capas viven como subpaquete directo bajo cada módulo y los ports cruzan el árbol solo por inyección en el composition root.

```text
access2web-blueprint/
├── platform/                          # NUEVO — monolito modular hexagonal (D68)
│   ├── pyproject.toml                 # pins: ruff==0.15.21, mypy==1.13.0, argon2-cffi==25.1.0
│   ├── Dockerfile                     # python:3.12-slim-bookworm@sha256:<digest>
│   ├── docker-compose.yml             # postgres + minio + backend dev (D77)
│   ├── .python-version                # 3.12
│   ├── Makefile                       # lint, typecheck, test, check-layers, security
│   ├── src/
│   │   ├── main.py                    # FastAPI app composition root
│   │   ├── shared/                    # cross-cutting (cache port, audit port, lockout)
│   │   │   ├── domain/
│   │   │   ├── ports/
│   │   │   ├── application/
│   │   │   └── adapters/
│   │   └── modules/
│   │       └── lanzadera/             # módulo admin (este change)
│   │           ├── domain/            # entities, value objects, invariants
│   │           │   ├── user.py        # users/spec.md
│   │           │   ├── app.py         # apps/spec.md
│   │           │   ├── profile.py     # profiles/spec.md
│   │           │   ├── assignment.py  # assignments/spec.md
│   │           │   ├── reset_token.py # auth-reset/spec.md
│   │           │   ├── global_admin.py# global_admins/spec.md
│   │           │   └── audit_event.py # audit/spec.md
│   │           ├── ports/             # Protocolos (interfaces)
│   │           │   ├── user_repository.py
│   │           │   ├── app_repository.py
│   │           │   ├── profile_repository.py
│   │           │   ├── assignment_repository.py
│   │           │   ├── reset_token_repository.py
│   │           │   ├── global_admin_repository.py
│   │           │   ├── credential_hasher.py          # Argon2id (D88)
│   │           │   ├── notification_delivery.py      # cola-por-tabla (D13)
│   │           │   ├── location.py                   # MVP assume_in_office
│   │           │   ├── cache.py                      # cachetools.TTLCache (D70)
│   │           │   ├── audit_log.py                  # audit/spec.md
│   │           │   ├── secret_manager.py             # key del DNI cifrado
│   │           │   └── bootstrap_admin_source.py     # GLOBAL_ADMIN_EMAILS (D91)
│   │           ├── application/       # use cases
│   │           │   ├── create_user.py
│   │           │   ├── disable_user.py
│   │           │   ├── grant_global_admin.py
│   │           │   ├── revoke_global_admin.py
│   │           │   ├── set_password.py               # CLI bootstrap (D91)
│   │           │   ├── issue_reset_token.py          # CRITICAL_HELPER (D90)
│   │           │   ├── consume_reset_token.py        # CRITICAL_HELPER (D90)
│   │           │   ├── assign_profile.py
│   │           │   ├── list_effective_apps.py
│   │           │   ├── audit_append.py
│   │           │   └── bootstrap_global_admins.py    # arranque idempotente
│   │           ├── adapters/          # driven — implementación
│   │           │   ├── persistence/    # seam único Postgres (DA-1)
│   │           │   │   ├── async_session_factory.py
│   │           │   │   └── repositories/
│   │           │   │       ├── user_repository_pg.py
│   │           │   │       ├── app_repository_pg.py
│   │           │   │       ├── profile_repository_pg.py
│   │           │   │       ├── assignment_repository_pg.py
│   │           │   │       ├── reset_token_repository_pg.py
│   │           │   │       ├── global_admin_repository_pg.py
│   │           │   │       ├── audit_log_pg.py
│   │           │   │       └── mail_queue_table_adapter.py
│   │           │   ├── crypto/
│   │           │   │   ├── credential_hasher_argon2id.py   # CRITICAL_HELPER (D88)
│   │           │   │   └── national_id_cipher.py          # AES-GCM vía secret manager
│   │           │   ├── notification/
│   │           │   │   └── mail_queue_table_adapter.py    # D13, cola-por-tabla
│   │           │   ├── location/
│   │           │   │   └── assume_in_office_adapter.py    # MVP H12
│   │           │   ├── cache/
│   │           │   │   └── ttl_cache_adapter.py           # cachetools.TTLCache (D70)
│   │           │   ├── audit/
│   │           │   │   └── structured_logger_adapter.py
│   │           │   ├── secrets/
│   │           │   │   └── env_secret_manager_adapter.py
│   │           │   └── bootstrap/
│   │           │       └── env_admin_source_adapter.py    # GLOBAL_ADMIN_EMAILS
│   │           ├── di/                # composition root del módulo
│   │           │   ├── container.py
│   │           │   └── bootstrap.py  # orquesta BootstrapAdapter al arranque
│   │           └── delivery/          # driving — FastAPI + HTMX + CLI
│   │               ├── http/
│   │               │   ├── app.py    # router FastAPI
│   │               │   ├── templates/async_jinja/
│   │               │   └── static/alpine.js
│   │               └── cli/
│   │                   └── platform_user.py     # set-password, grant/revoke
│   ├── migrations/                    # Alembic (D82 Expand and Contract)
│   │   ├── env.py
│   │   └── versions/
│   │       ├── 0001_core_schema.py
│   │       ├── 0002_seed_apps.py
│   │       ├── 0003_seed_profiles.py
│   │       ├── 0004_seed_users.py    # D89 — password_hash=NULL, status='password_reset_required'
│   │       ├── 0005_seed_assignments.py   # regla de mapeo legacy
│   │       └── 0006_seed_audit.py   # D55 — sin SSID/ubicación/coordenadas
│   ├── tests/
│   │   ├── test_ci_workflow.py       # pin del wiring QC-9
│   │   ├── test_gate_smoke.py        # QC-18 — smoke de cada gate
│   │   ├── test_no_legacy_in_auth.py # DA-13 — AST pinning
│   │   └── lanzadera/
│   │       ├── auth/test_hash_password.py
│   │       ├── auth/test_verify_password.py
│   │       ├── auth/test_issue_reset_token.py
│   │       ├── auth/test_consume_reset_token.py
│   │       ├── auth/test_no_legacy_compat.py
│   │       ├── assignments/test_sinacceso_exclusivity.py
│   │       ├── audit/test_same_transaction_audit.py
│   │       └── test_bootstrap_adapter.py
│   └── scripts/
│       ├── pytest_plugin/coverage_gate.py
│       ├── check_layers.py           # QC-2/QC-9
│       ├── check_complexity.py       # QC-1/QC-10
│       ├── check_pr_size.py          # QC-6
│       ├── check_branch_name.py      # QC-6
│       ├── check_crap.py             # QC-11
│       ├── check_dry.py              # QC-11
│       ├── quality_report.py         # QC-11 — agrega envelopes
│       └── migrate_from_access.py    # one-shot Dysflow → fixtures JSON
├── .github/workflows/
│   ├── ci.yml                        # todos los gates por PR
│   ├── security.yml                  # pip-audit, gitleaks, trivy config
│   ├── security-deep.yml             # semanal
│   └── pr-size.yml                   # check_pr_size.py
├── openspec/                         # EXISTENTE
└── docs/                             # EXISTENTE
```

El árbol se apega al gate `scripts/check_layers.py` con `ROOT_PACKAGE = "app.src.modules"`, `ALLOWED_IMPORTS` y `PURE_LAYERS` declarados en la sección §Decisiones arquitectónicas. Los archivos `domain/` no importan frameworks; los `adapters/` consumen drivers; el `delivery/` y `di/` actúan como composition root.

## Decisiones arquitectónicas del design

Las decisiones DA-* numeran el contrato que este `change` concreta. Cada fila declara la elección, la alternativa descartada y la decisión heredada (D-* o QC-*) que la respalda. Las nuevas formas no se inventan: cuando una decisión exige algo no presente en el material heredado, se marca explícitamente como `##ABIERTO##` en §Decisiones pendientes y gaps.

| ID | Decisión | Rationale | Alternativa descartada | Heredada |
|---|---|---|---|---|
| DA-1 | Capas `domain/`, `ports/`, `application/`, `adapters/`, `di/`, `delivery/` con `ALLOWED_IMPORTS` estricto y `PURE_LAYERS = {"domain", "ports", "application"}`; `ROOT_PACKAGE = "app.src.modules"` para `check_layers.py`; slicing vertical prohibido entre módulos salvo `shared` exento. | D8 exige hexagonal global; sin gate, la declaración queda en prosa y se viola sin huella (QC-9). | Sin gate: el hexagonal se «declara» y se viola silenciosamente (issue APAP_WEB #436). | D8, QC-2, QC-9 |
| DA-2 | Auth con Argon2id vía `argon2-cffi==25.1.0`, perfil `RFC_9106_LOW_MEMORY` (Argon2id, 64 MiB, 3 iteraciones, 4 hilos). `hash_password` y `verify_password` declarados `CRITICAL_HELPERS` al 100 % de cobertura (QC-5). | OWASP 2024 sitúa Argon2id como primera opción (memory-hard, resistente a GPU/ASIC). SHA256 sin salt del legacy no cumple estándares modernos (H1, resuelto). | bcrypt (`bcrypt==4.x`) — menos resistente a GPU que Argon2id; scrypt — sin bindings mantenidos en Python 3.12. | D88, D9, QC-5 |
| DA-3 | Schema `users.password_hash` NULL por defecto; `users.status` ENUM (`active`, `disabled`, `password_reset_required`, `locked`); **sin** columna `legacy_hash`, `password_legacy` ni `pass_hash_v1`. La columna `users.password_hash` se queda `NULL` para los 156 usuarios migrados; el flag `status='password_reset_required'` los bloquea hasta `consume_reset_token`. | D89 sustituye D36+D37 (obsoletos). Sin doble algoritmo, no hay superficie de ataque para un hash heredado sin sal. Migración 0004 no escribe hashes: el reset flow los crea (D88+D90). | Mantener `legacy_hash` + adapter de verificación legacy — rechazado por H1 (SHA256-hex sin sal no cumple OWASP 2024). | D89, D9 |
| DA-4 | Reset flow con tabla `reset_tokens` (`id`, `user_id`, `token_hash`, `expires_at`, `consumed_at`, `superseded_at`); `issue_reset_token(user_id) -> str` y `consume_reset_token(token, new_password) -> bool` atómicos y single-use; expiración 24 h; supersession al re-emitir; emisión del token crudo vía `NotificationDeliveryPort`. Ambos helpers declarados `CRITICAL_HELPERS`. | D90 exige tokens one-time de 24 h. La atomicidad previene reuso (race condition entre SELECT y UPDATE). El supersession anula el token anterior cuando el usuario pide otro. | Tokens JWT firmados — sin caducidad server-side fiable; sesiones persistentes con token persistente — no son one-time. | D90, D25, QC-5 |
| DA-5 | CLI `gentle-ai platform user set-password <email>` como camino **exclusivo** para crear el primer global admin antes de cualquier reset flow por email. El subcomando aplica Argon2id, marca `status='active'`, añade fila a `global_admins` y emite auditoría `auth.bootstrap.set_password`. | D91 cierra el bootstrap: el sistema no puede emitir tokens por email antes de tener un admin global que no haya pasado por email. | Reset flow auto-iniciado — rechazado por auth-reset (no_global_admin guard). | D25, D91 |
| DA-6 | `BootstrapAdapter` (driven) lee `GLOBAL_ADMIN_EMAILS` (semicolon-separated) en el arranque del proceso; siembra los `users` faltantes con `status='password_reset_required'`; crea filas en `global_admins`; emite auditoría `global_admins.bootstrap`. Idempotente: si el `user_id` ya existe y ya es global admin, no muta. Si la variable está unset, retorna sin error. | D21+D42+D48 aíslan la siembra de admins globales; el puerto `BootstrapAdminSource` permite sustituir la fuente (Vault, secret manager) sin tocar el composition root. | Inserción inline en el script de migración — rechazado por acoplamiento de secreto a DDL. | D21, D42, D48, D91 |
| DA-7 | Migraciones Alembic 0001-0006 con estrategia Expand and Contract (D82). La 0004 deja `password_hash = NULL` para los 156 usuarios (D89); no carga hashes legacy. La 0006 importa `TbConexiones` + `TbAplicacionesAperturas` sin SSID/ubicación/coordenadas (D55). Cada migración es aditiva; ninguna tira columnas legacy del backend autoritativo. | D82 prohíbe migraciones destructivas en una sola release. Plan de rollback del proposal: `DROP SCHEMA lanzadera CASCADE` deja el `.accdb` intacto. | Migración 0004 que invente hashes dummy o inserte `legacy_hash` — rechazado por DA-3. | D82, D55, D89 |
| DA-8 | `CachePort` con `cachetools.TTLCache` in-process como adapter MVP. Candidatos justificados: catálogo de `apps`, `profiles` por `app_id`, permisos efectivos del usuario. NO se cachean contadores de pendientes, métricas volátiles ni datos de sesión. TTL configurable por entrada; invalidación explícita al mutar. | D70 prohíbe caché de contadores y métricas volátiles; D71 deja Redis detrás del puerto como opción, no dependencia. | Redis desde día uno — rechazado por coste de operación para 200 usuarios concurrentes (D74). | D70, D71, D15 |
| DA-9 | `LocationPort` con implementación MVP `assume_in_office` (devuelve `True`). Refuerzo real cuando llegue el módulo HPS, que es la primera app office-only. | H12 (topología híbrida) marca el flag `apps.requires_office_presence` como no evaluado en el MVP. Implementar el adaptador real sin consumidor sería código especulativo (YAGNI). | Implementación con detección de IP corporativa — rechazado por ausencia de consumidor y por coupling a topología abierta (D75). | D53, H12 |
| DA-10 | Adapter de notificación v1 con cola-por-tabla: persiste filas en `mail_outbox` (`to`, `subject`, `body`, `status`, `created_at`); un dispatcher externo (futuro, no en este scope) las consume cada cinco minutos aprox. (legacy D65). El `NotificationDeliveryPort` define la firma `send(to, subject, body) -> None` para que el reset flow emita tokens sin acoplarse al canal. | D11+D13+D65 normalizan la notificación; v1 solo email; SMTP corporativo real sigue ABIERTO (P20). | Integración SMTP directa — rechazado por secreto IT aún no decidido. | D11, D12, D13, D65, P20 |
| DA-11 | Schema de auditoría sin columnas de telemetría (`ssid`, `bssid`, `coordinates`, `machine_name`, `ip_address`). El evento se persiste en la **misma transacción** que la mutación auth; si el insert de auditoría falla, la mutación hace rollback (DA-4 + audit/spec.md §Audit emission is mandatory). | D55 retira telemetría heredada. La atomicidad garantiza que un login sin audit es un login que no ocurrió. | Auditoría async (outbox + worker) — rechazado por ventana de inconsistencia entre mutación y evento. | D27, D28, D55, audit/spec.md |
| DA-12 | Mapeo legacy → profiles mediante tabla inmutable en código (`app/src/modules/lanzadera/domain/legacy_role_map.py`) que cubre los 7 flags del `TbUsuariosAplicacionesPermisos` (`Administrador`, `Calidad`, `CalidadAvisos`, `Técnico`, `Economía`, `Secretaría`, `SinAcceso`) más el rule `DEFAULT` para todas-NULL. Regla `SinAcceso` exclusivo implementada como cortocircuito: si el flag está activo, las demás filas se descartan. | D110 prohíbe el anti-patrón de campos dinámicos `F3..F9`. El mapping es declarativo, no inferido, y se testea con cardinalidad exhaustiva (H11). | Mapping por inferencia desde `TbUsuariosAplicaciones` — rechazado por anti-patrón de Password plano y por surface area de capacidades (D22). | D22, D45, D46, D110, H11 |
| DA-13 | Pin test AST que rechaza los símbolos `legacy_hash`, `verify_legacy`, `sha256`, `old_password`, `migrate_password` en `app/src/modules/lanzadera/auth/` y en todo `app/src/`. El test falla con `pytest_sessionfinish` mutando `session.exitstatus` si encuentra coincidencias. | D88+D89 sustituyen D36+D37; sin pinning, una reincorporación inadvertida del legacy vuelve a entrar (DA-3). | Confiar en revisión humana — rechazado por APAP_WEB #381 (gate ausente, drift reintroducido). | D88, D89, D9, QC-5 |

## Adaptadores y puertos concretos

Cada fila declara el puerto (interfaz `Protocol`), el adapter MVP que lo implementa y el test de contrato que verifica el comportamiento observable. Los nombres de archivo siguen el árbol de §Estructura física.

| Port | Interfaz (firma) | Adapter MVP | Test de contrato |
|---|---|---|---|
| `UserRepositoryPort` | `async def get_by_email(email: str) -> User | None`; `async def create(user: User) -> None`; `async def update_status(user_id: UUID, status: UserStatus) -> None`; `async def list_all() -> Sequence[User]`. | `UserRepositoryPg` sobre `asyncpg` + SQLAlchemy Core. | `tests/lanzadera/users/test_user_repository_contract.py` (in-memory fake). |
| `AppRepositoryPort` | `async def get_by_id(app_id: int) -> App | None`; `async def list_active() -> Sequence[App]`; `async def list_visible_to(user_id: UUID) -> Sequence[App]`. | `AppRepositoryPg`; cachea `list_active` vía `CachePort` (DA-8). | `tests/lanzadera/apps/test_app_repository_contract.py`. |
| `ProfileRepositoryPort` | `async def list_for_app(app_id: int) -> Sequence[Profile]`; `async def get_by_code(app_id: int, code: str) -> Profile | None`; `async def create(profile: Profile) -> None`; `async def set_active(profile_id: UUID, active: bool) -> None`. | `ProfileRepositoryPg`; cachea `list_for_app` con TTL 5 min (DA-8). | `tests/lanzadera/profiles/test_profile_repository_contract.py`. |
| `AssignmentRepositoryPort` | `async def create(user_id: UUID, app_id: int, profile_id: UUID) -> None`; `async def list_for_user(user_id: UUID) -> Sequence[Assignment]`; `async def list_for_app(app_id: int) -> Sequence[Assignment]`; `async def effective_permissions(user_id: UUID, app_id: int) -> Sequence[str]`. | `AssignmentRepositoryPg`; cachea `effective_permissions` por `(user_id, app_id)` con TTL 60 s e invalidación al mutar. | `tests/lanzadera/assignments/test_assignment_repository_contract.py`; incluye `test_sinacceso_exclusivity.py` (cardinalidad H11). |
| `CredentialHasherPort` | `def hash_password(plain: str) -> str`; `def verify_password(plain: str, hashed: str) -> bool`. | `CredentialHasherArgon2id` con `argon2-cffi==25.1.0` perfil `RFC_9106_LOW_MEMORY` (m=65536, t=3, p=4). | `tests/lanzadera/auth/test_hash_password.py` + `test_verify_password.py` (CRITICAL_HELPERS). |
| `ResetTokenRepositoryPort` | `async def insert(user_id: UUID, token_hash: str, expires_at: datetime) -> None`; `async def find_unused(token_hash: str) -> ResetToken | None`; `async def mark_consumed(token_hash: str, at: datetime) -> None`; `async def mark_superseded(user_id: UUID, at: datetime) -> None`; `async def purge_expired(now: datetime) -> int`. | `ResetTokenRepositoryPg`; unicidad por `token_hash`; índice en `(user_id, expires_at)`. | `tests/lanzadera/auth/test_issue_reset_token.py` + `test_consume_reset_token.py` (CRITICAL_HELPERS, atómicos). |
| `NotificationDeliveryPort` | `async def send(to: str, subject: str, body: str) -> None`. | `MailQueueTableAdapter`: inserta fila en `mail_outbox` con `status='pending'`. | `tests/lanzadera/notifications/test_mail_queue_contract.py`. |
| `LocationPort` | `async def is_user_in_office(user_id: UUID) -> bool`. | `AssumeInOfficeAdapter` devuelve `True` (DA-9, H12). | `tests/lanzadera/test_assume_in_office.py` (asserts explícito del stub). |
| `CachePort` | `def get(key: str) -> Any | None`; `def set(key: str, value: Any, ttl_seconds: int) -> None`; `def invalidate(prefix: str) -> int`. | `TtlCacheAdapter` con `cachetools.TTLCache(maxsize=1024)`. | `tests/lanzadera/test_ttl_cache_adapter.py`. |
| `AuditLogPort` | `async def append(event: AuditEvent) -> None`; `async def list_for_actor(actor_id: UUID, since: datetime) -> Sequence[AuditEvent]`. | `AuditLogPg`: insert atómico en la misma transacción que la mutación; `structured_logger_adapter` espeja a stdout canónico (D27). | `tests/lanzadera/audit/test_same_transaction_audit.py` (rollback si falla el insert). |
| `BootstrapAdminSource` (driven) | `def list_initial_emails() -> Sequence[str]`. | `EnvAdminSourceAdapter` lee `os.environ["GLOBAL_ADMIN_EMAILS"]` y parte por `;`; vacío si unset (DA-6). | `tests/lanzadera/test_bootstrap_adapter.py`. |
| `GlobalAdminRepositoryPort` | `async def list_all() -> Sequence[GlobalAdmin]`; `async def is_global_admin(user_id: UUID) -> bool`; `async def grant(user_id: UUID) -> None`; `async def revoke(user_id: UUID) -> None`. | `GlobalAdminRepositoryPg`; `revoke` rechaza si deja al sistema sin global admin (constraint en el repositorio). | `tests/lanzadera/global_admins/test_grant_revoke_round_trip.py` + `test_cannot_revoke_last.py`. |
| `SecretManagerPort` (cross-cutting) | `def get(key: str) -> str`. | `EnvSecretManagerAdapter` para MVP; producción intercambiable por Vault/AWS Secrets Manager sin tocar el dominio. | `tests/lanzadera/test_secret_manager_contract.py`. |

El `UserRepositoryPort.create` cifra `dni_encrypted` antes de persistir mediante `NationalIdCipher` (adapter de `SecretManagerPort`); la clave nunca aparece en logs ni en argumentos CLI (D25, D88-CA-S4). El contrato del `AuditLogPort.append` se invoca **dentro** de la sesión de SQLAlchemy iniciada por el caso de uso, de modo que un fallo del insert cause `ROLLBACK` de la mutación auth (audit/spec.md §Audit emission is mandatory + DA-11).

## Pipeline Alembic

Las seis migraciones siguen D82 (Expand and Contract). Cada fila declara qué crea, qué siembra o qué migra, y la decisión heredada que materializa. La 0004 deja `password_hash = NULL` para los 156 usuarios; la columna `legacy_hash` no existe.

| Versión | Crea / siembra / migra | Decisión heredada |
|---|---|---|
| `0001_core_schema` | Crea schema `lanzadera`; tablas `users`, `apps`, `profiles`, `user_app_assignments`, `global_admins`, `sessions`, `audit`, `reset_tokens`, `mail_outbox`; tipos ENUM `user_status`, `app_topology`, `app_registration_status`, `audit_event_type`; constraints de unicidad. | D8, D14, D27, D56, DA-1 |
| `0002_seed_apps` | Siembra las 20 filas de `apps` desde fixture `TbAplicaciones.json`; cada fila lleva `deployment_topology` y `requires_office_presence` derivados de `EjecucionEnOficina`; `registration_status='active'`. NO migra `Pass`, `Comando`, `URLDIrectorioIconoAplicacion` ni ningún atributo de lanzamiento. | D52, D53, D85, D58, apps/spec.md §Seed 20 apps |
| `0003_seed_profiles` | Para cada app, siembra al menos un profile con `code='default'`; además crea los profiles de mapeo legacy (`ADMIN`, `CALIDAD`, `CALIDAD_AVISOS`, `TECNICO`, `ECONOMIA`, `SECRETARIA`, `SIN_ACCESO`) cuando la app los tenía en `TbUsuariosAplicacionesPermisos`. `capabilities` es un JSONB mínimo inicial; el catálogo concreto de capabilities lo cierra `sdd-tasks` (§Decisiones pendientes — gap de profiles/spec.md). | D22, D45, D46, D110, profiles/spec.md §Standard profile |
| `0004_seed_users` | Migra los 156 usuarios desde fixture `tbUsuarios.json` con `email = lower(input.email)`, `name`, `dni_encrypted = encrypt(input.dni)`; **`password_hash = NULL`** y **`status = 'password_reset_required'`** para todos. **No** crea columna `legacy_hash`. Inserta auditoría `auth.bootstrap.set_password` para los usuarios creados vía CLI posteriormente. | D88, D89, DA-3, auth-core/spec.md §Seeded users start with reset required |
| `0005_seed_assignments` | Aplica la regla de mapeo legacy → `user_app_assignments`. Regla `SinAcceso` exclusivo: si `SinAcceso='Sí'`, solo se crea fila `SIN_ACCESO`. Sin `SinAcceso`: una fila por flag en `Sí` (`ADMIN`, `CALIDAD`, etc.). Todas-NULL → fila única `DEFAULT`. Tests de cardinalidad en `tests/lanzadera/assignments/test_sinacceso_exclusivity.py`. | D22, D42, D85, D102, D110, H11, assignments/spec.md |
| `0006_seed_audit` | Inserta eventos históricos desde fixtures `TbConexiones.json` y `TbAplicacionesAperturas.json`. **No** persiste SSID, BSSID, coordenadas, máquina ni IP. `TbConexiones` → `auth.login.success` / `auth.login.failure` según `Exitoso`. `TbAplicacionesAperturas` → `app.open` con `EnOficina` descartado. | D55, D27, D112, H4, audit/spec.md §Seed audit from legacy |

El orden 0001 → 0006 es **aditivo**: ninguna tira columnas ni borra filas en el schema ya creado. El rollback del MVP es `DROP SCHEMA lanzadera CASCADE;` — el `.accdb` legacy permanece intacto (proposal §Plan de rollback). Para un hotfix destructivo posterior (Contract), se sigue el flujo de dos releases según D82.

## Flujos críticos

Tres flujos concentran la superficie de seguridad y deben quedar ASCII-narrados antes del primer `git commit` de aplicación. Cada diagrama muestra el camino desde el driving adapter (CLI, web) hasta el driven adapter (Postgres, cola de notificación), pasando por el caso de uso y el puerto. Los nombres de archivo siguen el árbol de §Estructura física.

### Flujo A — Login con reset flow primero

El usuario migrado entra con `password_hash = NULL` y `status = 'password_reset_required'`. El primer login no puede autenticarlo; el sistema emite un token de reset por la cola de notificación. El segundo intento (con `consume_reset_token`) le asigna Argon2id y le concede acceso.

```text
[HTTP POST /login]
    └─→ application/login_user.py
         ├─→ ports/UserRepositoryPort.get_by_email  → adapters/postgres
         ├─ Si users.status == 'password_reset_required':
         │     └─→ application/issue_reset_token.py  (CRITICAL_HELPER, D90)
         │           ├─→ ports/ResetTokenRepositoryPort.insert
         │           ├─→ ports/NotificationDeliveryPort.send  → adapters/notification/mail_queue_table_adapter
         │           └─→ ports/AuditLogPort.append  (en la misma transacción)
         └─ En caso contrario:
               ├─→ ports/CredentialHasherPort.verify_password  (CRITICAL_HELPER, D88)
               ├─ Si ok → ports/AuditLogPort.append('auth.login.success')
               └─ Si nok → contador de fallos; al 5º, ports/UserRepositoryPort.update_status('locked') + audit 'auth.lockout' (D38-D40)
```

### Flujo B — Reset por enlace

El usuario llega al endpoint público `/reset` con el token crudo (string) recibido por email. La transacción consume el token, valida el hash Argon2id de la nueva contraseña y actualiza el estado del usuario. Si el insert de auditoría falla, la mutación rollbackea (DA-11).

```text
[HTTP POST /reset {token, new_password}]
    └─→ application/consume_reset_token.py  (CRITICAL_HELPER, D90)
         ├─→ ports/ResetTokenRepositoryPort.find_unused(token_hash)
         ├─ Si hash desconocido / expirado / superseded / consumido → return False
         ├─→ ports/CredentialHasherPort.hash_password(new_password)  (D88, RFC_9106_LOW_MEMORY)
         ├─→ ports/UserRepositoryPort.update(user_id, password_hash, status='active')
         ├─→ ports/ResetTokenRepositoryPort.mark_consumed(token_hash, now)
         └─→ ports/AuditLogPort.append('auth.reset.consumed')
              (todo en una sola transacción asyncpg; cualquier raise → ROLLBACK)
```

### Flujo C — Bootstrap del primer global admin

El operador arranca la plataforma por primera vez con `GLOBAL_ADMIN_EMAILS` definido. Antes de que ningún humano pueda usar `/login`, el operador ejecuta `gentle-ai platform user set-password <email>` para fijar la contraseña del primer admin y queda registrado como global admin. Solo entonces el flujo A puede emitir tokens por email.

```text
[Arranque del proceso]
    └─→ di/bootstrap.py
         └─→ application/bootstrap_global_admins.py
              ├─→ ports/BootstrapAdminSource.list_initial_emails()  → adapters/bootstrap/env_admin_source_adapter
              └─ Para cada email:
                    ├─ Si user no existe → ports/UserRepositoryPort.create(status='password_reset_required')
                    ├─→ ports/GlobalAdminRepositoryPort.grant(user_id)
                    └─→ ports/AuditLogPort.append('global_admins.bootstrap')

[CLI gentle-ai platform user set-password <email>]
    └─→ delivery/cli/platform_user.py
         └─→ application/set_password.py
              ├─→ ports/UserRepositoryPort.get_by_email
              ├─→ ports/CredentialHasherPort.hash_password(new_password)  (D88)
              ├─→ ports/UserRepositoryPort.update(password_hash, status='active')
              ├─→ ports/GlobalAdminRepositoryPort.grant(user_id)  (D91, primer admin)
              └─→ ports/AuditLogPort.append('auth.bootstrap.set_password')

    Ahora SÍ:
    └─→ cualquier /login con status='password_reset_required' puede emitir token por email
         (auth-reset/spec.md §no_global_admin guard desactivado)
```

## Modelo de datos

Las tablas principales del schema `lanzadera` derivan del contrato de los nueve sub-specs. La columna `users.legacy_hash` no existe (D89). La tabla `reset_tokens` es el hogar del reset flow (D90). La tabla `mail_outbox` materializa la cola-por-tabla del adapter v1 (D13). La tabla `audit` excluye telemetría (D55). Cada fila enumera las columnas operativas; los índices y constraints viven en la migración 0001.

| Tabla | Columnas operativas | Origen |
|---|---|---|
| `users` | `id` (UUID), `email` (citext, único, lower), `name`, `dni_encrypted` (bytea), `password_hash` (text NULL), `status` (ENUM `active`, `disabled`, `password_reset_required`, `locked`), `failed_attempts` (int default 0), `last_login_at` (timestamptz NULL), `created_at`, `updated_at`. | users/spec.md + D88+D89 |
| `apps` | `id` (int, rango baseline D85), `name`, `short_code` (único), `deployment_topology` (ENUM `central`, `office-nas`), `requires_office_presence` (bool), `registration_status` (ENUM `pending`, `active`, `retired`), `created_at`, `updated_at`. | apps/spec.md + D52+D58 |
| `profiles` | `id` (UUID), `app_id` (FK apps), `code` (text), `name`, `capabilities` (jsonb), `active` (bool default true), `created_at`, `updated_at`. UNIQUE (`app_id`, `code`). | profiles/spec.md + D22+D45 |
| `user_app_assignments` | `id` (UUID), `user_id` (FK), `app_id` (FK), `profile_id` (FK), `granted_by` (FK users, nullable), `granted_at`, `revoked_at` (NULL). UNIQUE (`user_id`, `app_id`, `profile_id`) cuando `revoked_at` IS NULL. | assignments/spec.md + D22+D42 |
| `global_admins` | `user_id` (FK users, PK). | global_admins/spec.md + D21 |
| `reset_tokens` | `id` (UUID), `user_id` (FK), `token_hash` (text, índice único), `expires_at` (timestamptz), `consumed_at` (NULL), `superseded_at` (NULL), `created_at`. | auth-reset/spec.md + D90 |
| `mail_outbox` | `id` (UUID), `to` (text), `subject` (text), `body` (text), `status` (ENUM `pending`, `sent`, `failed`), `attempts` (int default 0), `created_at`, `sent_at` (NULL). | D13, D65 |
| `audit` | `id` (UUID), `event_type` (text), `actor_id` (FK users, nullable), `target_id` (text, identificador opaco), `module` (text default `'lanzadera'`), `result` (text), `correlation_id` (UUID), `payload` (jsonb), `created_at`. **Sin** columnas `ssid`, `bssid`, `coordinates`, `machine_name`, `ip_address`. | audit/spec.md + D27+D55 |

El secret que cifra `users.dni_encrypted` vive en `SecretManagerPort` y se inyecta en el composition root. La clave nunca aparece en `os.environ` ni en argumentos CLI (CA-S4). El adapter `EnvSecretManagerAdapter` es el MVP; la sustitución por Vault/AWS Secrets Manager se hace cambiando el adapter sin tocar el dominio (D73, D11).

## Quality gates wiring

Los doce gates del MVP se commitean antes del primer `git commit` de código de aplicación. Cada uno aparece como step en `.github/workflows/ci.yml`, con su test de wiring (`tests/test_ci_workflow.py`) y su pin de versión.

| Gate | Mecanismo | Archivos | Decisiones |
|---|---|---|---|
| Lint base (`ruff check .`) | `pyproject.toml` `[tool.ruff]` con `select = ["E","F","W","I","UP","B"]`, pin exacto `ruff==0.15.21`. | `pyproject.toml`, `.github/workflows/ci.yml`. | QC-3 |
| Typecheck (`mypy app/src/`) | `pyproject.toml` `[tool.mypy]` con `python_version = "3.12"`, `enable_error_code = ["ignore-without-code"]`, `disallow_untyped_defs = true`. Pin `mypy==1.13.0`. | `pyproject.toml`. | QC-4 |
| Tests + cobertura (`pytest --cov=platform --cov-fail-under=85`) | `pyproject.toml` `[tool.coverage]` + `scripts/pytest_plugin/coverage_gate.py`. CRITICAL_HELPERS al 100 % vía `pytest_sessionfinish` mutando `session.exitstatus`. | `pyproject.toml`, `scripts/pytest_plugin/coverage_gate.py`, `.github/workflows/ci.yml`. | QC-5 |
| Hexagonal layer gate (`python scripts/check_layers.py`) | AST walk; `ROOT_PACKAGE = "app.src.modules"`; `ALLOWED_IMPORTS` y `PURE_LAYERS` declarados en DA-1; slicing vertical prohibido entre módulos salvo `shared` exento. | `scripts/check_layers.py`, `tests/lanzadera/test_layers_wiring.py`. | QC-2, QC-9, DA-1 |
| Complexity ceiling (`python scripts/check_complexity.py`) | AST + CC por función; **techo absoluto global `CC ≤ 15`** (QC-10). Nunca `top-N`. | `scripts/check_complexity.py`, `tests/lanzadera/test_complexity_wiring.py`. | QC-1 derivado, QC-10 |
| CRAP ceiling (`python scripts/check_crap.py`) | `CC² · (1 − cobertura)³ + CC` por función, techo `≤ 6` (QC-11). Consume `coverage.json`; falla cerrado si no existe. | `scripts/check_crap.py`. | QC-11 |
| DRY detector (`python scripts/check_dry.py`) | Clones type-1/type-2 sobre AST normalizado, 5+ sentencias, 0 tolerados. | `scripts/check_dry.py`. | QC-11 |
| Indicator aggregator (`python scripts/quality_report.py`) | Ejecuta layers → complexity → CRAP → DRY en orden fijo; emite `quality-report.json`; publica tabla en el job summary. | `scripts/quality_report.py`. | QC-11, regla 13 |
| PR size gate (`python scripts/check_pr_size.py`) | `git diff --stat` contra base; CRLF normalizado a LF; falla si `> 400`. Override `size:exception` exige justificación en cuerpo de PR. | `scripts/check_pr_size.py`, `.github/workflows/pr-size.yml`. | QC-6 |
| Branch-name gate (`python scripts/check_branch_name.py`) | Regex `^(feat\|fix\|refactor\|docs\|ci\|test)/<n>-<slug>$`; `main` allowlisted. | `scripts/check_branch_name.py`. | QC-6 |
| Secret scan (`gitleaks dir`) | `docker run zricethezav/gitleaks@sha256:<digest> dir . --redact --no-banner --exit-code 1`. | `.github/workflows/security.yml`. | QC-7 |
| Dependency scan (`pip-audit`) | Venv throwaway; `pip install -e .[dev]`; `pip-audit --skip-editable --strict`. | `.github/workflows/security.yml`. | QC-7 |
| Dockerfile scan (`trivy config`) | `docker run aquasec/trivy@sha256:<digest> config platform/Dockerfile --severity HIGH,CRITICAL`. | `.github/workflows/security.yml`. | QC-7 |
| Secret scan profundo + image scan | `gitleaks detect --source=.` con `fetch-depth: 0`; `trivy image` por digest declarado en `platform/Dockerfile`. | `.github/workflows/security-deep.yml`. | QC-7 |

Ningún step lleva `continue-on-error: true` ni `|| true` (Hard Rule 1 de `deterministic-quality-harness`). Las imágenes de los scanners van pinned por digest, nunca por tag (Hard Rule 6). El `Makefile` raíz expone `make lint`, `make typecheck`, `make test`, `make check-layers`, `make check-complexity`, `make check-crap`, `make check-dry`, `make check-pr-size`, `make check-branch-name`, `make security`, `make quality-report`.

## Tests y TDD

La disciplina es **strict TDD** (`openspec/config.yaml`: `apply.tdd: true`, `rules.apply.test_command: pytest --cov=platform --cov-fail-under=85`). El orden de adopción es el del plan por día de `docs/calidad-de-codigo-y-ci.md`. Los archivos siguientes son los que el MVP debe tener antes del primer `git commit` de código de `lanzadera.auth`.

| Archivo de test | Verifica | Decisión / Spec |
|---|---|---|
| `tests/test_ci_workflow.py` | Parsea `.github/workflows/ci.yml` y assertea la presencia de cada step de gate; cualquier改名 o borrado rompe el test. | QC-9, Hard Rule 4 (wiring pin). |
| `tests/lanzadera/auth/test_hash_password.py` | `hash_password` produce un PHC string Argon2id con `m=65536, t=3, p=4`; dos invocaciones con el mismo input producen hashes distintos (salt aleatorio); cobertura al 100 % del helper. | D88, DA-2, QC-5. |
| `tests/lanzadera/auth/test_verify_password.py` | `verify_password(plain, hash)` devuelve `True` para el par canónico y `False` ante cualquier perturbación; hash vacío o `None` lanzan `ValueError` (no se ejecuta Argon2id contra hash nulo). | D88, DA-2, QC-5. |
| `tests/lanzadera/auth/test_issue_reset_token.py` | `issue_reset_token(user_id)` persiste `token_hash`, `expires_at = now + 24h`, `consumed_at = NULL`, `superseded_at = NULL`; invoca `NotificationDeliveryPort.send` exactamente una vez; rechaza con `no_global_admin` cuando `global_admins` está vacío (salvo para `set-password` ya ejecutado). | D90, DA-4, QC-5. |
| `tests/lanzadera/auth/test_consume_reset_token.py` | `consume_reset_token(token, new_password)` actualiza `password_hash` y `status='active'` atómicamente; marca `consumed_at`; rechaza tokens desconocidos, expirados, superseded o ya consumidos; la transacción rollbackea si la auditoría falla. | D90, DA-4, DA-11, QC-5. |
| `tests/lanzadera/auth/test_no_legacy_compat.py` | AST walk sobre `app/src/modules/lanzadera/auth/` y `app/src/`; falla si encuentra los símbolos `legacy_hash`, `verify_legacy`, `sha256`, `old_password`, `migrate_password`. Se ejecuta como test pytest normal; la presencia de cualquiera de esos nombres falla el suite. | D88, D89, DA-13. |
| `tests/lanzadera/assignments/test_sinacceso_exclusivity.py` | Aplica la matriz legacy → `user_app_assignments` para los 32 casos (2^5 combinaciones de flags Sí/No × `SinAcceso` Sí/No); verifica la regla exclusiva de `SinAcceso`; verifica que `DEFAULT` aparece cuando todas son No/NULL; verifica que el resultado de la 0005 coincide con la cardinalidad esperada (≈622 filas). | D85, D102, H11, DA-12, assignments/spec.md. |
| `tests/lanzadera/audit/test_same_transaction_audit.py` | Mock del `AuditLogPort.append` que lanza `RuntimeError`; ejecuta el caso de uso `issue_reset_token`; verifica que `users.password_hash` y `reset_tokens.consumed_at` quedan sin mutar (rollback atómico). | D27, DA-11, audit/spec.md §Audit emission is mandatory. |
| `tests/lanzadera/test_bootstrap_adapter.py` | `BootstrapAdapter.bootstrap_global_admins()` con `GLOBAL_ADMIN_EMAILS='[email protected];[email protected]'`: crea dos users con `status='password_reset_required'`, dos filas en `global_admins`, una entrada de auditoría; segunda invocación no muta; con variable unset retorna sin error y sin filas. | D21, D42, D48, D91, DA-6, global_admins/spec.md §Bootstrap from GLOBAL_ADMIN_EMAILS. |
| `tests/lanzadera/test_layers_wiring.py` | Parsea `scripts/check_layers.py` y assertea `ROOT_PACKAGE == "app.src.modules"`; crea un fixture que viola `ALLOWED_IMPORTS` y assertea exit `1`; verifica `BASELINE` vacío. | QC-2, QC-9, DA-1. |
| `tests/lanzadera/test_no_legacy_in_auth.py` | Wrapper del `test_no_legacy_compat.py` con scope global a `app/src/` (defense-in-depth si se introduce un nuevo módulo con auth). | D88, DA-13. |
| `tests/test_gate_smoke.py` | Cada script `check_*.py` se invoca con un fixture que viola su contrato; se assertea exit `1` y se valida el envelope JSON. | Hard Rule 18, Execution Step 5. |

Los cuatro archivos `test_hash_password.py`, `test_verify_password.py`, `test_issue_reset_token.py`, `test_consume_reset_token.py` declaran `CRITICAL_HELPERS` y elevan la cobertura local al 100 % antes de admitir el helper en `coverage_gate.py` (QC-5). El orden de implementación sigue la curva RED → GREEN → REFACTOR; el refactor posterior nunca introduce duplicación detectable por `check_dry.py`.

## Riesgos de implementación

Cada riesgo declara severidad, mitigación y gate que la enforce. La tabla se cruza con la sección §Hallazgos y riesgos del proposal (H1-H14).

| ID | Riesgo | Mitigación | Gate |
|---|---|---|---|
| R-1 (H1) | El legacy almacena `Password` en texto plano (SHA256-hex sin sal confirmado por `integrations-security.md`); un adapter de verificación legacy reintroduciría un algoritmo de_hash sin sal. | D89 descarta hashes legacy; DA-3 prohíbe la columna `legacy_hash`; DA-13 pinea la ausencia vía AST. | `test_no_legacy_compat.py` + `coverage_gate` CRITICAL_HELPERS. |
| R-2 (H2) | Los binarios `.accdb` no viven en staging dentro del repo; `migrate_from_access.py` puede correr contra un `.accdb` distinto del baseline. | `migrate_from_access.py` documenta el SHA256 esperado del `.accdb` y aborta si no coincide; el comando vive en `Makefile` con `make migrate-fixtures`. | Smoke test del script en CI con fixture versionado. |
| R-3 (H3) | `codegraph-vba` no indexa Lanzadera en algunos worktrees (cross-project noise). | Cada worktree Lanzadera tiene su propio `.codegraph-vba/`; el MVP web no depende de la indexación VBA; el walkthrough v4 cubre 28/28 forms. | Verificación runtime al abrir worktree. |
| R-4 (H4) | Sincronización manual de datos comunes al NAS oficina (gap P-21). | Aceptado no bloqueante; apps office-only consultan snapshot; procedimiento operativo documentado en `docs/operacion/sync-nas-oficina.md` (pendiente `sdd-apply`). | Ninguno en MVP; H4 cierra en Fase 4. |
| R-5 (H5) | Topología de despliegue ABIERTA (D75). | Salvaguardada por D74 (no Kubernetes prematuro), D76 (PostgreSQL gestionado preferido). | Ninguno en MVP; H5 cierra cuando IT defina cloud. |
| R-6 (H6) | Adjuntos sin versionado + papelera 30 días (D17, D18). | No aplica al MVP de Lanzadera admin (no hay adjuntos en este slice). | Diferido al módulo de expedientes. |
| R-7 (H7) | Forms G4 con `MSComctlLib.TreeCtrl.2` + `ImageListCtrl.2` (3 forms). | Quedan en legacy archivado (D51); el admin web no requiere esa UI. | Walkthrough v4 documenta `skipped_tool_broken`. |
| R-8 (H9) | Booleanos como `Text(2)` (Sí/No) en 20+ columnas. | Migración 0004-0006 traduce `'Sí'` → `true`, `'No'` → `false`, otros → `NULL`. | Tests de cardinalidad sobre la matriz de mapping. |
| R-9 (H10) | `TbPermisos` con campos dinámicos `F3-F9`. | No se migra la estructura; las capabilities viven en `profiles.capabilities JSONB`. | DA-12 fija el shape. |
| R-10 (H11) | Regla `SinAcceso` exclusivo: si una fila legacy tiene `SinAcceso=Sí` y otros Sí, solo se crea `SIN_ACCESO`. | `legacy_role_map.py` declara el cortocircuito; `test_sinacceso_exclusivity.py` cubre los 32 casos. | DA-12 + test de cardinalidad. |
| R-11 (H12) | Topología híbrida: `apps.requires_office_presence` no se evalúa en el MVP. | `LocationPort` con `assume_in_office` que devuelve `True` (DA-9). El refuerzo llega con el módulo HPS. | Test `test_assume_in_office.py` assertea el stub. |
| R-12 (H13) | Diseño detallado del ciclo UAT ABIERTO (P16). | No bloquea el MVP; la gobernanza admin-global-only ya está APROBADA (D48). | Diferido a Fase 4. |
| R-13 (H14) | cosmic-ray DIFERIDO (QC-8). | Aceptado; se reactiva cuando ≥ 3 módulos tengan cobertura > 70 % y ≥ 5 tests por path real. | QC-8 cierra con disparador. |
| R-14 | Drift entre `docs/calidad-de-codigo-y-ci.md` y el código del MVP. | Hard Rule 10: cuando un doc diverge del código, el PR que detecta la divergencia corrige uno u otro. | `quality_report.py` + revisión de PR. |
| R-15 | Variabilidad de PYTHONHASHSEED, locale y orden de iteración afectan los envelopes de los gates. | `_pin_output_encoding()`, `hashlib.sha256` (no `hash()`), `sorted()` en cada walk, `-p no:randomly` en pytest. | Hard Rule 17 + tests de determinismo. |

## Plan de implementación por fases

Las cuatro fases corresponden al plan por día de `docs/calidad-de-codigo-y-ci.md` y a los tickets TK-LZ-MVP-1 a TK-LZ-MVP-24 del proposal. Cada fase termina con un PR ≤ 400 líneas (QC-6) y con `make quality-report` limpio.

### Fase 1 — Día 0 a 2: estructura + gates básicos + security scanning

Objetivo: dejar el repositorio listo para strict TDD sin código de aplicación todavía. Tickets: TK-LZ-MVP-1, TK-LZ-MVP-2, TK-LZ-MVP-3.

- Día 0 (0,5 h). Crear árbol `platform/`, `pyproject.toml`, `Dockerfile`, `docker-compose.yml`, `.python-version`, `.dockerignore`, `.gitignore` adicional, `Makefile`.
- Día 1 (4 h). Activar ruff + mypy + pytest + `coverage_gate` + PR size + branch name + build en `.github/workflows/ci.yml`. Pinear versiones. Wiring pins en `tests/test_ci_workflow.py`.
- Día 2 (3 h). Activar security scanning (`pip-audit`, `gitleaks`, `trivy`) con imágenes pinned por digest. `.github/workflows/security.yml` + `.github/workflows/security-deep.yml`. `.gitleaksignore` vacío al inicio.

Salida de fase: `make lint`, `make typecheck`, `make test`, `make security` ejecutan en CI y pasan sobre `platform/` vacío (smoke tests de los scripts).

### Fase 2 — Día 3 a 6: check_layers + check_complexity + primer módulo con TDD

Objetivo: gates hexagonales activos y `lanzadera.auth` implementado con strict TDD. Tickets: TK-LZ-MVP-4, TK-LZ-MVP-5, TK-LZ-MVP-7.

- Día 3 (8 h). Copia literal de `assets/scripts/check_layers.py` desde `deterministic-quality-harness`; ajustar `ROOT_PACKAGE = "app.src.modules"`; declarar `ALLOWED_IMPORTS` y `PURE_LAYERS` (DA-1). Wiring en CI. `tests/lanzadera/test_layers_wiring.py`.
- Día 4 (3 h). `check_complexity.py` con techo absoluto `CC ≤ 15`. Wiring. `tests/lanzadera/test_complexity_wiring.py`.
- Día 5 (1 h). Convenciones operativas en `AGENTS.md` raíz (Conf-1 a Conf-8).
- Día 6 (TDD). Primer módulo `app.src.modules.lanzadera.auth` con strict TDD: RED → GREEN → REFACTOR para `hash_password`, `verify_password`, `issue_reset_token`, `consume_reset_token`. Cobertura CRITICAL_HELPERS al 100 %. `test_no_legacy_compat.py` pinea la ausencia de legacy.

Salida de fase: `make quality-report` limpio; `coverage_gate` exige 100 % sobre los cuatro CRITICAL_HELPERS.

### Fase 3 — Apply: Alembic 0001-0006 + adapters + delivery + CLI

Objetivo: el sistema corre end-to-end en local con `docker-compose up`. Tickets: TK-LZ-MVP-8 a TK-LZ-MVP-22, TK-LZ-MVP-24.

- Aplicar 0001_core_schema, 0002_seed_apps, 0003_seed_profiles, 0004_seed_users, 0005_seed_assignments, 0006_seed_audit. Cada migración en su propio PR ≤ 400 líneas.
- Implementar adapters Postgres para los seis repositorios (DA-1). In-memory fakes para tests.
- Implementar `MailQueueTableAdapter` (DA-10) + `TtlCacheAdapter` (DA-8) + `AssumeInOfficeAdapter` (DA-9) + `EnvAdminSourceAdapter` + `BootstrapAdapter` (DA-6).
- Delivery HTMX: rutas FastAPI por sub-spec (users, apps, profiles, assignments, global_admins, audit). Templates Jinja async + Alpine.js (D67).
- CLI `gentle-ai platform user ...` con subcomandos `set-password`, `grant-global-admin`, `revoke-global-admin`, `list-apps`, `assign-profile` (D25, D26, D91).
- `scripts/migrate_from_access.py` one-shot Dysflow → fixtures JSON; ejecutable en local y en CI con `make migrate-fixtures` (DA-7).

Salida de fase: smoke E2E en local — login con `set-password`, asignación de profile, reset flow end-to-end.

### Fase 4 — UAT y cut-over

Objetivo: el ecosistema migra del `.accdb` al web. Tickets: TK-LZ-MVP-23.

- Runbook de UAT (P16) con gobernanza admin-global-only (D48).
- Cut-over one-shot desde el legacy; `DROP SCHEMA lanzadera CASCADE` como rollback documentado.
- Sincronización manual al NAS oficina (H4) según `docs/operacion/sync-nas-oficina.md`.
- Cierre del change: archive `lanzadera-mvp` con deltas de specs.

Salida de fase: `openspec/changes/lanzadera-mvp/archive/` con `archive.md` + specs actualizadas.

## Decisiones pendientes y gaps

Este diseño cierra los huecos que le corresponden y deja abiertos los que `sdd-tasks` o fases posteriores deben resolver. Cada item declara la decisión heredada de origen, si se cierra aquí, y el camino si queda abierto.

| # | Gap | Origen | Estado en este design |
|---|---|---|---|
| G-1 | Política exacta de normalización del email (lowercase completo vs `local-part` lower + `domain` lower, IDN de segundo nivel). | users/spec.md L84 (`##ABIERTO##`) | Cerrado en DA-3 con `email = lower(input.email)`; IDN queda como mejora futura (no se modela dominio IDN). |
| G-2 | Set canónico de capabilities por app para los 20 IDs en alcance. | profiles/spec.md L83 (`##ABIERTO##`) | ABIERTO — `sdd-tasks` debe producir `app/src/modules/lanzadera/domain/legacy_role_map.py` con `capabilities` mínimos por código (`ADMIN`, `CALIDAD`, etc.). El `profiles.capabilities` se siembra en 0003 con un JSONB vacío o con un shape provisional. |
| G-3 | Lista de campos de catálogo legacy que NO migran (`Pass`, `Comando`, `URLDIrectorioIconoAplicacion`, etc.). | apps/spec.md L82 (`##ABIERTO##`) | Cerrado en DA-7 (migración 0002): solo migran `id`, `name`, `short_code`, `deployment_topology`, `requires_office_presence`, `registration_status`. La lista completa de campos descartados se documenta en `docs/03-aplicaciones/lanzadera/data-model.md` durante `sdd-apply`. |
| G-4 | Política de auditoría para intentos fallidos de crear admin global (severidad, notificación a SOC). | auth-bootstrap/spec.md L37 (`##ABIERTO##`) | ABIERTO — el MVP registra el evento en el log canónico con severidad `WARN`; el contrato exacto de notificación a SOC queda para cuando la organización defina su canal de seguridad (no bloqueante). |
| G-5 | Política exacta del canal de notificación al usuario cuando se emite un token (P20, SMTP corporativo real). | auth-reset/spec.md L62 (`##ABIERTO##`) | ABIERTO — el MVP usa la cola por tabla como adapter v1 (DA-10); el contrato del `NotificationDeliveryPort` queda fijo (`send(to, subject, body) -> None`). La sustitución por SMTP corporativo se hace reemplazando el adapter sin tocar dominio. |
| G-6 | Política de expiración de contraseña configurable (D41). | auth-core/spec.md L79 (`##ABIERTO##`) | ABIERTO — el MVP implementa `status='password_reset_required'` sin caducidad periódica (DA-3 + DA-4); el reset flow cubre la recuperación. La caducidad configurable se introduce en una release posterior cuando se defina la política de cumplimiento. |
| G-7 | Peso de `SinAcceso` en la navegación del menú global. | assignments/spec.md L86 (ABIERTO) | ABIERTO — el menú global filtra apps según `effective_permissions`; el comportamiento exacto de `SIN_ACCESO` (mostrar app deshabilitada vs ocultar) se decide en la épica de UI de `sdd-apply`. El adapter de menú se declara en `app/src/modules/lanzadera/application/` y queda stub. |
| G-8 | Periodos definitivos de retención por cumplimiento normativo (P10). | audit/spec.md L85 (ABIERTO) | ABIERTO — el MVP aplica 90 días hot + 1 año total (D29 provisional). El puerto `AuditRetentionPort` permite cambiar la política sin tocar el dominio. |

Gaps cerrados por este diseño: G-1 (DA-3), G-3 (DA-7). G-2, G-4, G-5, G-6, G-7, G-8 quedan abiertos para `sdd-tasks` o releases posteriores; ninguno bloquea la entrada a la fase de aplicación. Ningún gap abierto requiere nueva decisión D-* ni nueva regla QC-*: se resuelven con tareas acotadas que el sub-agente `sdd-tasks` debe planificar.

### Blockers para `sdd-tasks`

`Sdd-tasks` puede arrancar sin bloqueos. Los elementos que requiere son:

1. Confirmación de que `pyproject.toml` declara los pins exactos (`argon2-cffi==25.1.0`, `cachetools==5.x`, `sqlalchemy==2.0.x`, `alembic==1.13+`, `asyncpg==0.30+`, `fastapi==0.119+`, `jinja2==3.1+`, `htmx==2.0.4`, `alpinejs==3.15+`).
2. Disponibilidad de la fixture `TbAplicaciones.json`, `tbUsuarios.json`, `TbUsuariosAplicacionesPermisos.json`, `TbConexiones.json`, `TbAplicacionesAperturas.json` con SHA256 documentado, producidas por `scripts/migrate_from_access.py` desde `C:\00repos\datos\Lanzadera_Datos.accdb` (H2, R-2).
3. Cierre del G-2 (capacidades mínimas por `profile_code`) por el equipo de producto antes de la migración 0003, o aceptación de sembrar `capabilities={}` provisional y dejar la canonicalización para una release posterior.
4. Confirmación del environment variable `GLOBAL_ADMIN_EMAILS` con al menos un email válido antes del primer arranque del MVP (DA-5 + DA-6 + D91).

Si alguno de estos puntos falta, `sdd-tasks` debe marcarlo como `##ABIERTO##` en el primer ticket del lote, no avanzar en silencio.

## Contratos críticos (firma exacta de los `CRITICAL_HELPERS`)

Los cuatro helpers declarados `CRITICAL_HELPERS` se escriben en `app/src/modules/lanzadera/application/` como funciones puras o casos de uso. Las firmas siguen los contratos de los sub-specs. El coverage gate exige 100 % de líneas ejecutadas sobre cada uno (QC-5). El pin test `test_no_legacy_compat.py` rechaza cualquier import que reintroduzca SHA256 u `old_password` en el árbol.

```python
# app/src/modules/lanzadera/ports/credential_hasher.py
from __future__ import annotations
from typing import Protocol


class CredentialHasherPort(Protocol):
    """Argon2id via argon2-cffi==25.1.0 profile RFC_9106_LOW_MEMORY (D88)."""

    def hash_password(self, plain: str) -> str: ...
    def verify_password(self, plain: str, hashed: str) -> bool: ...
```

```python
# app/src/modules/lanzadera/application/issue_reset_token.py
from __future__ import annotations
from datetime import datetime, timedelta, timezone
import secrets
from argon2 import PasswordHasher
from argon2.exceptions import VerifyMismatchError


PHC_PARAMS = {"memory_cost": 65536, "time_cost": 3, "parallelism": 4, "hash_len": 32, "salt_len": 16}


def _hasher() -> PasswordHasher:
    return PasswordHasher(**PHC_PARAMS)


def hash_password(plain: str) -> str:
    if not plain:
        raise ValueError("plain must be non-empty")
    return _hasher().hash(plain)


def verify_password(plain: str, hashed: str) -> bool:
    if not hashed:
        return False
    try:
        return _hasher().verify(hashed, plain)
    except VerifyMismatchError:
        return False
```

```python
# app/src/modules/lanzadera/application/consume_reset_token.py
from __future__ import annotations
import hashlib
from datetime import datetime, timezone
from uuid import UUID

from app.src.modules.lanzadera.ports.audit_log import AuditLogPort
from app.src.modules.lanzadera.ports.credential_hasher import hash_password
from app.src.modules.lanzadera.ports.reset_token_repository import ResetTokenRepositoryPort
from app.src.modules.lanzadera.ports.user_repository import UserRepositoryPort


def _sha256(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


async def issue_reset_token(
    user_id: UUID,
    *,
    reset_tokens: ResetTokenRepositoryPort,
    notifications: NotificationDeliveryPort,
    global_admins: GlobalAdminRepositoryPort,
    audit: AuditLogPort,
) -> str:
    """Issue a one-time token; supersedes prior unconsumed tokens. D90."""
    if not await global_admins.is_there_any() and not await _first_admin_already_set(user_id, audit):
        raise NoGlobalAdminError("cannot issue reset before bootstrap")
    raw_token = secrets.token_urlsafe(32)
    token_hash = _sha256(raw_token)
    now = datetime.now(timezone.utc)
    await reset_tokens.mark_superseded_for(user_id=user_id, at=now)
    await reset_tokens.insert(user_id=user_id, token_hash=token_hash, expires_at=now + timedelta(hours=24))
    await notifications.send(to=..., subject="Reset", body=f"Use this token: {raw_token}")
    await audit.append(event_type="auth.reset.issued", actor_id=user_id)
    return raw_token


async def consume_reset_token(
    token: str,
    new_password: str,
    *,
    reset_tokens: ResetTokenRepositoryPort,
    users: UserRepositoryPort,
    audit: AuditLogPort,
) -> bool:
    """Atomic consume. False on unknown / expired / superseded / already-used."""
    token_hash = _sha256(token)
    record = await reset_tokens.find_unused(token_hash=token_hash, now=datetime.now(timezone.utc))
    if record is None:
        return False
    new_hash = hash_password(new_password)
    await users.set_password_and_activate(user_id=record.user_id, password_hash=new_hash)
    await reset_tokens.mark_consumed(token_hash=token_hash, at=datetime.now(timezone.utc))
    await audit.append(event_type="auth.reset.consumed", actor_id=record.user_id)
    return True
```

Las firmas aquí declaradas son el contrato que `sdd-tasks` debe implementar test por test. Cualquier desviación (por ejemplo, cambiar `expires_at = now + 24h` por un valor distinto) debe re-validarse contra el sub-spec y contra este diseño.

## Runbook de migración

La ejecución local y en CI sigue el mismo orden. Cada comando es idempotente en su tramo: re-correr no rompe datos ya migrados, pero aborta ante una desviación del SHA256 de la fixture.

```bash
# 1. Generar fixtures desde el .accdb (sólo local con Dysflow; CI consume fixtures versionadas)
python scripts/migrate_from_access.py \
    --accdb "C:/00repos/datos/Lanzadera_Datos.accdb" \
    --out platform/tests/fixtures/ \
    --expected-sha256 "<registrado en docs/03-aplicaciones/lanzadera/data-model.md>"

# 2. Levantar Postgres + MinIO
docker compose up -d postgres

# 3. Aplicar 0001-0006 en orden
alembic upgrade head

# 4. Ejecutar suite
pytest --cov=platform --cov-fail-under=85

# 5. Smoke del gate hexagonal
python scripts/check_layers.py --root .

# 6. Reporte de indicadores
python scripts/quality_report.py --root . --out quality-report.json
```

El rollback completo (sólo MVP, sin release en producción) es:

```bash
docker compose down
psql -U platform -d platform -c "DROP SCHEMA IF EXISTS lanzadera CASCADE;"
```

El `.accdb` legacy permanece intacto y operativo. Para un Contract posterior a una release caliente, se siguen dos releases según D82: la release N añade la nueva columna con default, la N+1 retira la vieja cuando el monitor confirma 0 uso.

## Operación del MVP

Las variables de entorno, los puertos listening y los entry points del proceso se documentan en una sola tabla para que la rampa de operación no disperse la información entre `docker-compose.yml`, `pyproject.toml` y los scripts. La disciplina sigue la convención Conf-5 a Conf-8 de `docs/calidad-de-codigo-y-ci.md`.

### Variables de entorno

| Variable | Tipo | Default MVP | Lectura | Decisión |
|---|---|---|---|---|
| `DATABASE_URL` | str | `postgresql+asyncpg://platform:platform@localhost:5432/platform` | composition root (Postgres). | D14 |
| `PLATFORM_SECRET_KEY` | str (Fernet) | generada en primer arranque si falta, persistida en volumen `secrets/` | `SecretManagerPort.get("users.dni.cipher")`. | CA-S2, D25 |
| `GLOBAL_ADMIN_EMAILS` | str (semicolon-separated) | unset (bootstrap no hace nada) | `EnvAdminSourceAdapter.list_initial_emails()`. | DA-6, D91 |
| `LOCKOUT_THRESHOLD` | int | `5` | `LockoutPolicy`. | D38 |
| `LOCKOUT_DURATION_SECONDS` | int | `3600` | `LockoutPolicy`. | D39 |
| `RESET_TOKEN_TTL_SECONDS` | int | `86400` | `issue_reset_token`. | D90 |
| `CACHE_TTL_APPS_LIST` | int | `300` | `TtlCacheAdapter` para `AppRepository.list_active`. | DA-8 |
| `CACHE_TTL_PROFILES_BY_APP` | int | `300` | `TtlCacheAdapter` para `ProfileRepository.list_for_app`. | DA-8 |
| `CACHE_TTL_EFFECTIVE_PERMS` | int | `60` | `TtlCacheAdapter` para `AssignmentRepository.effective_permissions`. | DA-8 |
| `LOG_LEVEL` | str | `INFO` | `structured_logger_adapter`. | D27 |
| `OTEL_EXPORTER_OTLP_ENDPOINT` | str | unset | deshabilitado en MVP;预留 cuando se decida observabilidad distribuida. | D27, P14 |

Ninguna variable contiene secretos sin cifrar. `PLATFORM_SECRET_KEY` se inyecta vía secret manager real (Vault, AWS Secrets Manager) en producción y se materializa en el archivo de secretos del contenedor en dev (CA-S2, CA-S4).

### Puertos y entry points

| Entry point | Tipo | Descripción |
|---|---|---|
| `app.src.main:app` | HTTP | FastAPI app; expone `GET /`, `POST /login`, `POST /logout`, `POST /reset`, rutas admin bajo `/admin/...`. Sirve HTMX fragments con `Cache-Control` fingerprint. |
| `app.src.main:bootstrap` | lifecycle | Se ejecuta al arrancar el proceso; llama al `BootstrapAdapter` (DA-6) y al `MailQueueTableAdapter.purge_expired` en background. |
| `gentle-ai platform user ...` | CLI | Subcomandos: `set-password`, `grant-global-admin`, `revoke-global-admin`, `list-apps`, `assign-profile`. Acceso restringido a global admin (D25); secretos vía variables de entorno (CA-S4). |
| `alembic upgrade head` | migration | Aplica 0001-0006; cada migración en su propio down-grade explícito. |
| `python scripts/migrate_from_access.py` | one-shot | Genera fixtures JSON desde el `.accdb` con Dysflow read-only (R-2). |

### Quick map inverso del módulo

| Si necesita | Abra primero | Y luego consulte |
|---|---|---|
| Entender qué hace cada sub-spec | `openspec/changes/lanzadera-mvp/specs/<sub>/spec.md`. | §Adaptadores y puertos concretos. |
| Ver la forma hexagonal exigida | §Estructura física del módulo. | `scripts/check_layers.py` con `ALLOWED_IMPORTS` (DA-1). |
| Implementar un helper crítico | §Contratos críticos. | `tests/lanzadera/auth/test_<helper>.py`. |
| Aplicar las migraciones | §Runbook de migración. | `docs/08-decisiones-y-preguntas-abiertas.md` (D82, D89). |
| Operar el MVP en local | §Operación del MVP (variables, puertos, entry points). | `docker-compose.yml`, `Makefile`. |
| Diagnosticar por qué un gate falla | §Quality gates wiring. | `docs/calidad-de-codigo-y-ci.md` §Anti-patrones. |
| Auditar la ausencia de legacy | `tests/lanzadera/auth/test_no_legacy_compat.py`. | §Riesgos de implementación (R-1). |
| Cerrar un gap abierto | §Decisiones pendientes y gaps. | El sub-spec de origen (referenciado en la columna «Origen»). |

## Convenciones operativas aplicadas al módulo

Las ocho convenciones de `docs/calidad-de-codigo-y-ci.md` §Convenciones operativas se aterrizan aquí para el módulo `lanzadera`. La tabla no duplica la definición del gate; solo nombra cómo aplica a este módulo concreto.

| Conf | Convención | Aplicación en `lanzadera` |
|---|---|---|
| Conf-1 | Layer boundaries respetan `ALLOWED_IMPORTS` y `PURE_LAYERS`. | `domain/` no importa `argon2`, `sqlalchemy`, `asyncpg`, `fastapi`; `ports/` solo importa `domain/` y `ports/`; `application/` no importa `adapters/` (verificado por `check_layers.py`). |
| Conf-2 | `Annotated[T, Depends(get_x)]` para dependencias FastAPI. | `delivery/http/app.py` declara las dependencias con `Annotated[UserRepositoryPort, Depends(get_user_repository)]`. |
| Conf-3 | `# type: ignore[<código>]` con código explícito, nunca bare. | `pyproject.toml` declara `enable_error_code = ["ignore-without-code"]`; cualquier `# type: ignore` sin código rompe el typecheck. |
| Conf-4 | `# pragma: no cover` con razón en la línea adyacente. | Reservado a `_pin_output_encoding` y adapters de logging cuyo fallo no es de código de aplicación; cada pragma lleva `  # pragma: no cover  # rationale: <razón>`. |
| Conf-5 | `from __future__ import annotations` en cada `.py`. | Cada archivo del módulo y de los tests empieza con esa línea; lo verifica `ruff format --check`. |
| Conf-6 | PR ≤ 400 líneas; nombre de rama `^(feat\|fix\|refactor\|docs\|ci\|test)/<n>-<slug>$`. | `check_pr_size.py` mide `git diff --stat` con CRLF normalizado a LF; `check_branch_name.py` aplica la regex. Override `size:exception` exige justificación en el cuerpo de PR (Hard Rule 11 — segregación de roles). |
| Conf-7 | Migraciones siempre backward-compatibles con estrategia Expand and Contract. | 0001-0006 son aditivas; ninguna tira columnas legacy. El Contract exige dos releases según D82. |
| Conf-8 | `coverage_gate.py` marca los `CRITICAL_HELPERS` del módulo. | `CRITICAL_HELPERS = ["hash_password", "verify_password", "issue_reset_token", "consume_reset_token"]` declarados en `scripts/pytest_plugin/coverage_gate.py`. La función `pytest_sessionfinish` muta `session.exitstatus` (Hard Rule 8). |

## Referencias

| Recurso | Ruta |
|---|---|
| Propuesta del change | `openspec/changes/lanzadera-mvp/proposal.md` |
| Sub-specs | `openspec/changes/lanzadera-mvp/specs/{users,apps,profiles,assignments,auth-core,auth-reset,auth-bootstrap,global_admins,audit}/spec.md` |
| Decisiones D5-D91 | `docs/08-decisiones-y-preguntas-abiertas.md` |
| Arquitectura objetivo y principios | `docs/09-arquitectura-objetivo-y-principios.md` |
| Quality gates del MVP | `docs/calidad-de-codigo-y-ci.md` |
| Auditoría Clean Code + swarm-forge | `docs/auditoria-harnesses-clean-code.md` |
| Configuración del change | `openspec/config.yaml` |
| Skill de quality harness | `.opencode/skills/deterministic-quality-harness/SKILL.md` (symlinked desde `~/.config/opencode/skills/`) |
| Skill de documentación | `~/.config/opencode/skills/documentation-alan-style/` |

## Lista de comprobación final

- [ ] Cada DA-* referencia al menos una D-* o QC-* heredada; ningún DA-* inventa una decisión arquitectónica nueva (DA-13 cierra G-1 y G-3; DA-7 cierra la lista de campos legacy no migrados).
- [ ] Los seis gaps del sub-agente `sdd-spec` (G-1, G-2, G-3, G-4, G-5, G-6) están clasificados como cerrados o abiertos con destino claro.
- [ ] Los dos gaps adicionales detectados (G-7 peso de `SinAcceso` en menú, G-8 retención definitiva) están registrados.
- [ ] Las seis migraciones Alembic 0001-0006 siguen Expand and Contract (D82); la 0004 deja `password_hash = NULL`; la 0006 omite SSID/ubicación/coordenadas.
- [ ] El gate hexagonal `check_layers.py` está dimensionado con `ROOT_PACKAGE = "app.src.modules"`, `ALLOWED_IMPORTS` y `PURE_LAYERS` declarados.
- [ ] Los cuatro `CRITICAL_HELPERS` (`hash_password`, `verify_password`, `issue_reset_token`, `consume_reset_token`) están listados con cobertura al 100 % (QC-5).
- [ ] Los doce quality gates aparecen wired en `ci.yml`, pinned por `tests/test_ci_workflow.py` y commiteados antes del primer `git commit` de código de aplicación.
- [ ] El tono es Castellano peninsular formal en el cuerpo narrativo; inglés en nombres de archivo, código y secciones técnicas.
- [ ] No se crean archivos `tasks.md` ni `archive.md` en este `change` (los abren `sdd-tasks` y `sdd-archive`).

## Siguiente paso

`Sdd-tasks` lee este diseño y produce `openspec/changes/lanzadera-mvp/tasks.md` con la descomposición de tickets, agrupados por fase y con cada ticket completable en una sesión (regla `rules.tasks` del config). El contrato de capabilities entre `sdd-design` y `sdd-tasks` vive en §Estructura física del módulo y §Adaptadores y puertos concretos.