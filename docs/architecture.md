[← Back to CODEBASE-GUIDE](../../CODEBASE-GUIDE.md)

# access2web-blueprint — Architecture

**Esta guía es la fuente de verdad única para la arquitectura del monorepo. Complementa a `CODEBASE-GUIDE.md` (overview + ownership) y a `DOCS.md` (technical reference).** Aquí vive la decisión técnica: las capas hexagonales, los puertos y adapters, las decisiones D-<n> cross-cutting, las DA-<n> del MVP, los gaps conocidos y el path de lectura inversa.

> **Sentence que organiza el repo entero**: «Lanzadera es la madre: ahí nacen usuarios, aplicativos y permisos. Las otras 7 apps son consumidoras.»

> **Architectural sentence**: «La hexagonalidad no se discute, se verifica. Cada decisión arquitectónica aprobada se traduce en una forma física, un puerto y un gate que falla cuando el código la viola.» — `openspec/changes/lanzadera-mvp/design.md`

## What this is / is not

| Es | No es |
|---|---|
| Catálogo de las decisiones D-<n> y DA-<n> vigentes, con estado (`vigente`, `OBSOLETO`). | Réplica de CODEBASE-GUIDE.md ni de DOCS.md. |
| Mapa de los patrones hexagonales (capas, puertos, composition root). | Manual de uso de cada `check_*.py`. |
| Estado actual del stack y sus versiones pinned. | Roadmap de lo que viene. |
| Gaps conocidos entre las docs y el código real. | Plan para cerrar esos gaps (eso vive en issues). |

## 90-second mental model

```text
                    ┌──────────────────────────────┐
                    │  Legacy Access/VBA (.accdb)   │
                    │  en data/staging/<app>/ (R2)  │
                    └──────────────┬───────────────┘
                                   │ dysflow + codegraph-vba
                                   ▼
┌──────────────────────────────────────────────────────────────────┐
│  Este monorepo                                                     │
│                                                                   │
│  docs/03-aplicaciones/<app>/{epic,walkthrough-*.json}             │
│  ├── Lanzadera ★ (mother)                                         │
│  ├── Expedientes, HPS, HPS_Solicitudes, Brass,                     │
│  │   Gestion_Riesgos, NoConformidades, Condor                      │
│                                                                   │
│  app/  (Python hexagonal, renombrado desde platform/ 2026-08-09)  │
│  ├── src/modules/<app>/{domain,ports,application,adapters,        │
│  │                       di,delivery}                              │
│  ├── src/shared/  (cross-cutting: cache, audit, lockout)          │
│  ├── migrations/versions/   (Alembic; Expand and Contract, D82)    │
│  ├── tests/   (TDD + Playwright E2E; gates por módulo)            │
│  └── pytest_plugin/   (coverage_gate, etc.)                       │
│                                                                   │
│  openspec/changes/<change>/{proposal,design,specs/*,tasks}.md    │
│  scripts/check_*.py   (12 quality gates, ver §CI)                  │
│  skills/    (copy de las skills internalizadas)                    │
└──────────────────────────────────────────────────────────────────┘
                                   │
                                   ▼
                ┌──────────────────────────────────────┐
                │ PostgreSQL (esquema por módulo, D14) │
                │ S3-compatible (anexos, D16)           │
                │ + Redis (F2, detrás de CachePort)    │
                │ + Mistica + HTMX + Alpine.js         │
                └──────────────────────────────────────┘
```

## Las 8 aplicaciones

| App | Rol | Walkthrough | Epic | Código en `app/src/modules/` |
|---|---|---|---|---|
| **lanzadera** | Madre (users + apps + permissions) | v4 mergeada | merged (PR #3) | MVP — pendiente de extracción a `app/src/modules/lanzadera/` |
| **expedientes** | Consumer + producer E2E | v4 mergeada | merged (PR #4) | plan PR-08 |
| hps, hps-solicitudes, brass, gestion-riesgos, no-conformidades, condor | Consumers | mergeada (research cerrado) | mergeada | pendiente |

## Stack

| Capa | Versión pinned | Decisión |
|---|---|---|
| Python | 3.12+ | D66 |
| Backend | FastAPI 0.119+, SQLAlchemy 2.0.x, Alembic 1.13+, asyncpg 0.30+ | D67 |
| Frontend | HTMX 2.0.4, Jinja2 3.1+, Alpine.js 3.15+ | D67 |
| DB | PostgreSQL, esquema por módulo | D14 |
| Auth crypto | argon2-cffi 25.1.0, perfil RFC_9106_LOW_MEMORY | DA-2, D88 |
| Caché | cachetools.TTLCache (in-process MVP) | DA-8, D70 |
| Object storage | S3-compatible (no AWS-only) | D16 |
| UI tokens | Mistica design system | — |

## Layout del repo (real, post-rename 2026-08-09)

```text
access2web-blueprint/
├── README.md, CONTRIBUTING.md, AGENTS.md, CODEBASE-GUIDE.md
├── DOCS.md, CHANGELOG.md
├── docs/
│   ├── architecture.md           ← este fichero
│   ├── 03-aplicaciones/<app>/   ← epic, walkthrough, capabilities per app
│   ├── design/mockups/          ← UI Mistica autocontenida
│   └── prompts/                 ← reportes a mantenedor dysflow
├── app/                         ← hexagonal Python (renombrado desde platform/)
│   ├── README.md                ← naming deviation justificada
│   ├── pyproject.toml           ← ruff==0.15.21, mypy==1.13.0, argon2-cffi==25.1.0
│   ├── Dockerfile               ← python:3.12-slim-bookworm
│   ├── docker-compose.yml       ← postgres + minio + dev
│   ├── src/
│   │   ├── main.py              ← FastAPI composition root
│   │   ├── shared/              ← cross-cutting (cache, audit, lockout)
│   │   └── modules/<app>/{domain,ports,application,adapters,di,delivery}
│   ├── migrations/versions/     ← Alembic 0001..0006 (D82)
│   ├── tests/                   ← TDD + Playwright + `test_ci_workflow.py` (QC-9)
│   └── pytest_plugin/           ← coverage_gate, etc.
├── openspec/changes/<change>/{proposal,design,specs/*,tasks}.md
├── scripts/check_*.py           ← 12 quality gates (ver §CI gates)
├── .github/workflows/{ci,security,security-deep,codeql,release}.yml
├── data/staging/                ← binarios legacy (R2-pulled)
└── skills/                      ← copy internalizada de skills del proyecto
```

## Decisiones arquitectónicas D-<n> cross-cutting vigentes

Las decisiones se numeran con D-<n> por orden de aprobación, no secuencial. Las obsoletas se reemplazan por una decisión nueva; las duplicadas se descartan. Esta tabla es **discontinua por diseño**.

| ID | Decisión | Aplica a | Estado |
|---|---|---|---|
| D5 | Plataforma modular permission-aware | Todas | vigente |
| D8 | Hexagonal global | Todas | vigente |
| D14 | Esquema por módulo en PostgreSQL | Todas | vigente |
| D16 | Object storage S3-compatible | NC, GR, Brass, Lanzadera | vigente |
| D21, D42, D48 | Bootstrap de admins globales aislado | Lanzadera | vigente |
| D27, D28 | Logs estructurados (no tablas TbLog) | Cross-cutting | vigente |
| D36+D37 | Hash heredado sin sal | — | **OBSOLETO** — reemplazado por D88+D89 |
| D52, D53 | Topología de las apps | Todas | vigente |
| D55 | Sin telemetría (SSID/ubicación/coordenadas) | Cross-cutting | vigente |
| D65 | Notificación cola-por-tabla | Lanzadera | vigente |
| D66-D68 | Stack (Python/FastAPI) + monolito modular | Todas | vigente |
| D70 | No caché de contadores ni métricas | Cross-cutting | vigente |
| D71 | Redis detrás del CachePort, no dependencia | Cross-cutting | vigente |
| D77 | Docker desde día uno | Plataforma | vigente |
| D82 | Expand and Contract en migraciones | Todas | vigente |
| D85 | Catálogo de 20 aplicaciones | Lanzadera | vigente |
| D88 | Argon2id vía argon2-cffi==25.1.0 | Auth | vigente |
| D89 | Sin columna legacy_hash; `password_hash=NULL` para migrados | Auth | vigente |
| D90 | Reset flow con tokens one-time 24 h | Auth | vigente |
| D91 | CLI exclusivo para primer admin global | Lanzadera | vigente |
| D110 | Anti-patrón campos dinámicos F3..F9 | Lanzadera | vigente |
| D112 | TbAplicacionesAperturas: limpieza de telemetría | Lanzadera | vigente |

## Decisiones de diseño del MVP (DA-<n>)

Las DA-<n> extienden las D-<n> dentro del change `openspec/changes/lanzadera-mvp/`. Promueven a D-<n> cuando el change se archive, o se descartan si se reemplazan. Esta tabla es **puente**, no duplicado; la fuente vive en `design.md` del change.

| ID | Resumen |
|---|---|
| DA-1 | Capas hexagonales con `ROOT_PACKAGE = "app.src.modules"` (post-rename), `PURE_LAYERS = {domain, ports, application}`. |
| DA-2 | Argon2id perfil RFC_9106_LOW_MEMORY. QC-5 exige 100 % en `CredentialHasherArgon2id.hash` y `.verify`. |
| DA-3 | `users.password_hash` NULL + `status` ENUM. Sin columna legacy. |
| DA-4 | Reset tokens atómicos, expiración 24 h, supersession. |
| DA-5 | CLI `gentle-ai platform user set-password` exclusivo para bootstrap. |
| DA-6 | `BootstrapAdapter` idempotente desde `GLOBAL_ADMIN_EMAILS`. |
| DA-7 | Alembic 0001..0006 aditivas. Rollback = `DROP SCHEMA lanzadera CASCADE;` (legacy intacto). |
| DA-8 | `CachePort` con `TTLCache`. NO se cachean counters ni métricas. |
| DA-9 | `LocationPort` con `assume_in_office` MVP (stub). |
| DA-10 | Notificación cola-por-tabla; SMTP real abierto. |
| DA-11 | Audit en **misma transacción** que la mutación auth. |
| DA-12 | Mapping legacy → profiles en tabla inmutable, regla `SinAcceso` exclusivo. |
| DA-13 | Pin test AST rechaza símbolos `legacy_hash`, `sha256`, `migrate_password`. |

## Patrón hexagonal (DA-1)

| Capa | Pureza | Puede importar | Lo prohíbe `scripts/check_layers.py` |
|---|---|---|---|
| `domain/` | PURE | Sólo el propio módulo | frameworks, ports, adapters |
| `ports/` | PURE | `typing.Protocol` | implementaciones concretas, frameworks |
| `application/` | PURE | domain + ports | adapters, delivery |
| `adapters/` | driven | domain, ports, frameworks | delivery |
| `di/` | composition root | todos | — |
| `delivery/` | driving | application, ports, adapters | domain |
| `shared/` (cross-cutting) | mixto | cache, audit, lockout | slicing vertical entre módulos |

`scripts/check_layers.py` con `ROOT_PACKAGE = "app.src.modules"` enforza esta tabla. Tests en `tests/test_ci_workflow.py` pinean el wiring contra drift.

## Auth (D88-D91 + DA-2..DA-5)

- **Crypto**: Argon2id perfil `RFC_9106_LOW_MEMORY`. QC-5 resuelve `CredentialHasherArgon2id.hash` y `.verify` sobre la clase y exige 100 % de cobertura.
- **Bootstrap**: `gentle-ai platform user set-password <email>` es el único camino para crear el primer admin. `BootstrapAdapter` siembra los siguientes vía `GLOBAL_ADMIN_EMAILS`. Idempotente.
- **Reset**: tokens one-time, 24 h, atómicos. `issue_reset_token` y `consume_reset_token` ambos `CRITICAL_HELPERS`.
- **Audit**: `auth.bootstrap.set_password`, `auth.login.success`, `auth.login.failure` van en **misma transacción** que la mutación (DA-11).

## Persistence (D14, D82)

- Una base de datos PostgreSQL compartida. Cada módulo tiene su propio schema (`lanzadera`, próximos: `expedientes`, etc.).
- Alembic con `expand_and_contract` (D82): cada release es aditiva. Rollback = `DROP SCHEMA <módulo> CASCADE;` (legacy intacto).
- Migraciones `0001..0006` son la línea base del MVP. Cada futura app tiene su propia serie numerada.

### Adapters Postgres (DA-1, W01..W06)

- Ruta canónica de los adapters: `app/src/modules/lanzadera/adapters/persistence/repositories/<entity>_pg.py`. La ruta legacy `adapters/repos/` se conserva como stub vacío para enlazar imports durante la transición.
- `async_session_factory(url)` construye el seam único: `AsyncEngine` + `AsyncSessionFactoryPort` con `search_path = lanzadera, public`. Todos los adapters toman ese puerto por constructor; ninguno importa `sqlalchemy.engine` directamente.
- Adapters implementados en el seam: `UserRepositoryPg`, `AppRepositoryPg`, `ProfileRepositoryPg`, `AssignmentRepositoryPg`, `AuditLogPg`, `GlobalAdminRepositoryPg`, `ResetTokenRepositoryPg`, `MailQueueTableAdapter`. Cada uno declara su propio `sa.Table` para mantener la tabla-reflection consigo mismo.
- `AsyncSessionFactory` expone dos context-manager helpers:
  - `read_only_session()`: yields una `AsyncSession` y la cierra al exit, sin commit. Para paths de lectura.
  - `transaction()`: yields una `AsyncSession`, commitea al exit, rollback en exception. Para paths de escritura.
- Cada adapter abre y cierra su `AsyncSession` por método. El boilerplate `try/except/finally` queda cerrado desde W08..W20 (8 PRs de prelude-cleave consecutivos).

## Caché (D70-D71 + DA-8)

- `CachePort` con `TTLCache` in-process (MVP).
- **NO** se cachean: contadores de pendientes, métricas volátiles, datos de sesión.
- Redis queda detrás del puerto como opción futura (F2 del MVP) — no es dependencia.

## Audit (D27, D55 + DA-11)

- Logs estructurados; nunca se persiste `ssid`, `bssid`, `coordinates`, `machine_name`, `ip_address` (D55).
- Audit en **misma transacción** que la mutación: si el insert falla, la mutación hace rollback (DA-11).
- Eventos canónicos: `auth.bootstrap.set_password`, `auth.login.success`, `auth.login.failure`, `app.open`, `global_admins.bootstrap`.

## Secret manager (D9-D10, D25)

- `SecretManagerPort` con `EnvSecretManagerAdapter` en MVP.
- Producción intercambiable por Vault o AWS Secrets Manager sin tocar el dominio.
- Prohibido en logs ni en argumentos CLI.

## Notificaciones (D11-D13, D65 + DA-10)

- `NotificationDeliveryPort` con cola-por-tabla: filas en `mail_outbox` con `status='pending'`.
- Dispatcher externo (fuera de scope MVP) las consume cada ~5 min.
- SMTP real queda abierto hasta que IT confirme el secreto corporativo.

## Migración de binarios

- Staging de `.accdb` vía R2 bucket `access2web-staging-binaries`.
- Scripts: `scripts/setup-staging.ps1` (inicial) y `scripts/sync-to-r2.ps1` (diario incremental).
- Lanzadera es **source-only** en este repo (D156): el walkthrough audit se hace contra el pull de R2.

## CI gates

Trece `scripts/check_*.py` + `.github/workflows/`:

| Gate | Vive en | QC |
|---|---|---|
| `check_branch_name.py` | scripts/ | QC-6 |
| `check_pr_size.py` | scripts/ | QC-6 |
| `check_layers.py` | scripts/ | QC-2, QC-9 |
| `check_complexity.py` | scripts/ | QC-1, QC-10 |
| `check_dry.py` | scripts/ | QC-11 |
| `check_decision_guards.py` | scripts/ | Decisiones críticas |
| `check_legacy_hashes.py` | scripts/ | DA-13 (pin AST) |
| `check_legacy_retirement.py` | scripts/ | Retirada de rutas legacy |
| `check_mutation_sites.py` | scripts/ | mutation semanal |
| `check_mutation.py` | scripts/ | mutation semanal |
| `check_test_classification.py` | scripts/ | Clasificación de tests |
| `check_workflows.py` | scripts/ | Integridad YAML, pins, concurrencia y aislamiento de PRs públicos |
| `quality_report.py` | scripts/ | QC-11 (agregador) |
| `check_walkthrough_schema.py` | scripts/ | MUST fields del template `walkthrough.json` (53 walkthroughs) |
| `pytest_plugin/coverage_gate.py` | app/ | QC-5 |
| ci.yml | .github/workflows/ | orquesta todo por PR |
| security.yml | .github/workflows/ | pip-audit, gitleaks, trivy config |
| security-deep.yml | .github/workflows/ | semanal |
| codeql.yml | .github/workflows/ | cada PR, push a main y semanal |
| release.yml | .github/workflows/ | tags + identidad + firma keyless del digest |

## Gaps conocidos (lo que la doc dice y el código no)

Estos gaps están aquí hasta que se cierren. No se disimulan; se documentan para que la siguiente IA no se confunda.

| Gap | Impacto | Acción propuesta |
|---|---|---|
| `CODEBASE-GUIDE.md` y `DOCS.md` aún dicen `platform/`, pero la realidad es `app/` desde 2026-08-09. | IAs que leen CODEBASE-GUIDE antes de actuar rompen el layout actual. | Editar ambos docs para usar `app/` consistentemente. PR aparte. |
| `openspec/changes/lanzadera-mvp/design.md` se escribió antes del rename — sigue diciendo `platform.src.modules`. | Drift entre design histórico y código actual. | PR a design.md para apuntar a `app.src.modules`. Mantener el change como histórico. |
| `inputs/` existe en el repo pero no aparece en CODEBASE-GUIDE ni en DOCS. | IAs no saben qué va ahí. | Documentar en §Layout o eliminar si quedó vacío. |
| D-<n> numeración no es contigua (D36+D37 obsoletos, huecos varios). | IA que busca «todos los D» queda descolocada. | Esta tabla los cataloga con estado `vigente` u `OBSOLETO`; mantenerla viva. |

## Recommended reading path

| Step | Doc | Read this when… |
|---|---|---|
| 1 | [README](../../README.md) | Necesita entender el producto en 5 min. |
| 2 | [DOCS](../../DOCS.md) | Quiere el technical reference raíz. |
| 3 | [CODEBASE-GUIDE](../../CODEBASE-GUIDE.md) | Es mantenedor o contribuidor nuevo. |
| 4 | **Esta guía** | Va a tocar código de plataforma o necesita una decisión arquitectónica. |
| 5 | `openspec/changes/<change>/design.md` | Está implementando un change específico. |
| 6 | `docs/03-aplicaciones/<app>/epic.md` | Necesita entender una app legacy concreta. |

## Quick map: if you need X, open Y

| If you need to… | Open first | Then check |
|---|---|---|
| Validar que un cambio respeta las capas hexagonales | Esta guía §Patrón hexagonal | `scripts/check_layers.py` |
| Crear un nuevo módulo (app) | Esta guía §Layout + DA-1 | `openspec/changes/lanzadera-mvp/design.md` §Estructura física |
| Tocar autenticación, bootstrap o reset | Esta guía §Auth + DA-2..DA-5 | `lanzadera-mvp/design.md` DA-2..DA-5 |
| Modificar la migración Alembic | Esta guía §Persistence + D82 | `lanzadera-mvp/design.md` DA-7 |
| Añadir un puerto o adapter cross-cutting | Esta guía §Caché / §Audit / §Secret / §Notificaciones | `lanzadera-mvp/design.md` §Adaptadores y puertos |
| Auditar un cambio que afecta varias apps | Esta guía §Decisiones D-<n> cross-cutting | las `epic.md` de las apps afectadas |

## What this is NOT

- **No** es el lugar para introducir decisiones nuevas — eso es un PR al `design.md` del change correspondiente (vía SDD).
- **No** documenta el plan de release — eso es [CHANGELOG](../../CHANGELOG.md).
- **No** es el manual de uso de los `check_*.py` — cada script tiene su docstring; este doc mapea cuál aplica a qué decisión.
- **No** reemplaza a `CODEBASE-GUIDE.md` (overview + ownership) ni a `DOCS.md` (technical reference).

## Core invariants

- **Hexagonalidad verificable, no discutida**: cada capa (`domain/`, `ports/`, `application/`, `adapters/`, `di/`, `delivery/`) tiene reglas de importación que `scripts/check_layers.py` enforza con `ROOT_PACKAGE = "app.src.modules"`. Tests en `tests/test_ci_workflow.py` pinean `ROOT_PACKAGE`, `ALLOWED_IMPORTS` y `PURE_LAYERS` contra drift.
- **D-<n> vigentes son vinculantes**: las decisiones de la tabla §Decisiones arquitectónicas D-<n> con estado `vigente` son contrato. Modificarlas exige PR al `design.md` del change correspondiente (vía SDD) o un issue que las reemplace por una decisión nueva; el doc nunca se edita a mano para «relajar» una regla.
- **Expand and Contract en cada release (D82)**: toda migración Alembic es aditiva. Rollback = `DROP SCHEMA <módulo> CASCADE;` con el legacy intacto. Las migraciones 0001..0006 son la línea base del MVP; cada futura app abre su propia serie.
- **Audit en la misma transacción que la mutación auth (DA-11)**: los eventos canónicos (`auth.bootstrap.set_password`, `auth.login.success`, `auth.login.failure`, `app.open`, `global_admins.bootstrap`) se persisten atómicamente con la mutación. Si el insert de audit falla, la mutación hace rollback.
- **Sin telemetría sensible (D55)**: nunca se persiste `ssid`, `bssid`, `coordinates`, `machine_name`, `ip_address` en logs estructurados ni en argumentos CLI. El check AST `scripts/check_legacy_hashes.py` pinea este invariante.
- **Pin de hashes legacy eliminado (D88+D89)**: la columna `legacy_hash` no existe; los passwords migrados van con `password_hash = NULL`. `scripts/check_legacy_hashes.py` rechaza los símbolos `legacy_hash`, `sha256`, `migrate_password` en código nuevo (DA-13).

## Contributor checklist

- [ ] La decisión propuesta (D-<n> nueva) se discute primero en un PR al `openspec/changes/<change>/design.md`, no se añade directamente a la tabla §Decisiones arquitectónicas.
- [ ] Si la decisión cruza varias apps, se marca como `cross-cutting` y se referencia desde la `epic.md` correspondiente.
- [ ] El cambio de código respeta las 6 reglas de §Core invariants; el `ci / quality` check pasa verde.
- [ ] El pin AST de `check_legacy_hashes.py` se mantiene verde (no se reintroducen `legacy_hash`, `sha256` ni `migrate_password`).
- [ ] Si la decisión toca un puerto (`CachePort`, `AuditLogPort`, `SecretManagerPort`, `LocationPort`, `NotificationDeliveryPort`), el adapter mantiene la signatura del puerto y el dominio no importa el framework concreto.
- [ ] Si la decisión introduce una migración Alembic, sigue `expand_and_contract` (D82): añadir columnas nullable o tablas nuevas, sin `DROP` ni `ALTER` destructivos en la misma release.
- [ ] El PR es ≤ 400 líneas (`additions + deletions`); si no, partir por unidad de trabajo o encadenar.
- [ ] Se actualiza la tabla §Decisiones arquitectónicas D-<n> y §Gaps conocidos si corresponde.

## Navigation

Previous: [CODEBASE-GUIDE](../../CODEBASE-GUIDE.md) | Next: [calidad-de-codigo-y-ci](calidad-de-codigo-y-ci.md)

---

[Next: CHANGELOG →](../../CHANGELOG.md)
