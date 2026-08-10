# Tasks: Lanzadera MVP — Primer slice de la plataforma hexagonal

> Change: `lanzadera-mvp` · Project: `access2web-blueprint` · strict_tdd: true · pytest: `pytest --cov=app --cov-fail-under=85`

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Total WUs | 36 |
| Max lines / WU | 320 |
| Avg lines / WU | ~100 |
| 400-line budget risk | **Low** |
| Chained PRs recommended | **Yes** |
| Delivery strategy | auto-chain |
| Chain strategy | stacked-to-main |

```
Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: stacked-to-main
400-line budget risk: Low
```

### Suggested Work Units

| WU | Conventional-commit title | Vertical slice | Est. lines | Focused test | Runtime harness | Rollback boundary | Base |
|----|--------------------------|----------------|------------|--------------|----------------|-------------------|------|
| F1 | `feat(lanzadera): scaffold platform directory tree` | Empty `app/` tree with `__init__.py` stubs for all future packages; establishes physical layout | 80 | N/A | N/A | Delete `app/src/modules/lanzadera/` | main |
| F2 | `feat(ci): add pyproject.toml with exact dependency pins` | All pins declared (argon2-cffi, sqlalchemy, alembic, fastapi, etc.); ruff/mypy configs | 120 | `ruff check app/` | N/A | Revert file | main |
| F3 | `feat(infra): add Dockerfile and docker-compose.yml` | Multi-stage Dockerfile (python:3.12-slim-bookworm@sha256:...); compose with postgres + minio; PLATFORM_SECRET_KEY in secrets/ volume | 140 | `docker compose up -d postgres` | `docker compose ps` | `docker compose down` | main |
| F4 | `feat(ci): add Makefile with all quality targets` | Targets: lint, typecheck, test, check-layers, check-complexity, check-crap, check-dry, check-pr-size, check-branch-name, security, quality-report, migrate-fixtures | 100 | `make lint` | N/A | Revert file | main |
| F5 | `feat(ci): add GitHub Actions CI workflow` | `.github/workflows/ci.yml` with ruff/mypy/pytest/coverage_gate/steps; all gates pinned; `tests/test_ci_workflow.py` verifies gate wiring | 160 | `pytest tests/test_ci_workflow.py` | N/A | `git revert HEAD` | main |
| F6 | `feat(security): add security scanning workflows` | `pip-audit` (venv throwaway), `gitleaks` (digest pinned), `trivy config` (digest pinned); weekly deep scan workflow | 140 | `gitleaks detect --source=. --redact` | N/A | `git revert HEAD` | main |
| F7 | `feat(qa): add hexagonal layer gate scripts` | `check_layers.py` from deterministic-quality-harness/assets; `ALLOWED_IMPORTS`, `PURE_LAYERS` per DA-1; `test_layers_wiring.py` | 160 | `python app/scripts/check_layers.py --root .` | N/A | Revert script | main |
| F8 | `feat(qa): add complexity, CRAP, DRY and quality-report scripts` | `check_complexity.py` (CC≤15), `check_crap.py` (CRAP≤6), `check_dry.py` (0 clones), `quality_report.py` (fixed order); wiring tests | 200 | `python app/scripts/check_complexity.py --root .` | N/A | Revert scripts | main |
| F9 | `feat(qa): add PR-size and branch-name gate scripts` | `check_pr_size.py` (400-line gate), `check_branch_name.py` (regex), `coverage_gate.py` declaring `CRITICAL_HELPERS`; smoke test | 120 | `python app/scripts/check_pr_size.py --root .` | N/A | Revert scripts | main |
| D1 | `feat(domain): add User, Session, ResetToken entities` | Pure domain: `User` (id, email, name, dni_encrypted, password_hash, status ENUM, failed_attempts, last_login_at), `Session`, `ResetToken` VO; no framework imports | 160 | `pytest tests/lanzadera/domain/` | N/A | Delete domain files | main |
| D2 | `feat(domain): add App, Profile, Assignment entities` | Pure domain: `App` (topology ENUMs, registration_status), `Profile` (capabilities JSONB), `Assignment` (user×app×profile); no framework imports | 140 | `pytest tests/lanzadera/domain/` | N/A | Delete domain files | main |
| D3 | `feat(domain): add GlobalAdmin, AuditEvent, legacy_role_map` | Pure domain: `GlobalAdmin` VO, `AuditEvent` (no telemetry per DA-11), `legacy_role_map.py` (7 legacy flags → profile codes, SinAcceso cortocircuito, 32-case matrix) | 180 | `pytest tests/lanzadera/domain/` | N/A | Delete domain files | main |
| D4 | `feat(domain): add LockoutPolicy value object` | Pure domain: `LockoutPolicy(threshold, duration_seconds)`; `Session` entity with expiry | 80 | `pytest tests/lanzadera/domain/` | N/A | Delete domain files | main |
| M1 | `chore(migrations): add 0001_core_schema` | Creates `lanzadera` schema; tables: users, apps, profiles, user_app_assignments, global_admins, sessions, audit, reset_tokens, mail_outbox; ENUMs; unique constraints; indexes; **no** telemetry columns | 220 | `alembic upgrade head` | Fresh `initdb` | `alembic downgrade -1` | main |
| M2 | `chore(migrations): add 0002_seed_apps` | Seed 20 apps from `TbAplicaciones.json`; maps `EjecucionEnOficina` → topology/requires_office; excludes Pass/Comando/URLDIrectorioIconoAplicacion (DA-7); RED test first | 140 | `alembic upgrade head && pytest tests/lanzadera/migrations/` | Same compose | `alembic downgrade -1` | main |
| M3 | `chore(migrations): add 0003_seed_profiles` | Seed `default` profile per app; legacy-role profiles (ADMIN, CALIDAD, etc.) when legacy had them; `capabilities={}` (G-2 ABIERTO); RED test first | 120 | `alembic upgrade head && pytest tests/lanzadera/migrations/` | Same compose | `alembic downgrade -1` | main |
| M4 | `chore(migrations): add 0004_seed_users` | Migrate 156 users with `email=lower(input.email)`, `dni_encrypted=encrypt(input.dni)`; `password_hash=NULL`, `status='password_reset_required'`; NO legacy_hash column (DA-3); RED test first | 160 | `alembic upgrade head && pytest tests/lanzadera/migrations/` | Same compose | `alembic downgrade -1` | main |
| M5 | `chore(migrations): add 0005_seed_assignments` | Apply legacy-role mapping (legacy_role_map.py); SinAcceso exclusive cortocircuito; DEFAULT for all-NULL; cardinality ≈622 rows (H11); RED test first | 140 | `alembic upgrade head && pytest tests/lanzadera/assignments/` | Same compose | `alembic downgrade -1` | main |
| M6 | `chore(migrations): add 0006_seed_audit` | Seed TbConexiones → auth.login.success/failure; seed TbAplicacionesAperturas → app.open; NO SSID/BS/coords/machine/IP (D55); RED test first | 120 | `alembic upgrade head && pytest tests/lanzadera/audit/` | Same compose | `alembic downgrade -1` | main |
| C1 | `feat(migrate): add migrate_from_access.py one-shot script` | Dysflow read-only extract → JSON fixtures; SHA256 verified against `.accdb`; produces TbAplicaciones.json, tbUsuarios.json, TbUsuariosAplicacionesPermisos.json, TbConexiones.json, TbAplicacionesAperturas.json; callable via `make migrate-fixtures` | 280 | `python app/scripts/migrate_from_access.py --help` | N/A (local/CI only) | Revert script | main |
| A1 | `feat(auth): add hash_password and verify_password (CRITICAL_HELPERS)` | Argon2id via `argon2-cffi==25.1.0` RFC_9106_LOW_MEMORY; `hash_password` and `verify_password` in `application/credential_helpers.py`; RED tests first; 100% coverage gate | 180 | `pytest tests/lanzadera/auth/test_hash_password.py tests/lanzadera/auth/test_verify_password.py --cov=app --cov-fail-under=85` | N/A | Revert helper files | main |
| A2 | `feat(auth): add issue_reset_token and consume_reset_token (CRITICAL_HELPERS)` | `issue_reset_token` (24h, supersedes prior, NoGlobalAdminError guard); `consume_reset_token` (atomic, single-use, rollback on audit failure); RED tests first; 100% coverage gate | 280 | `pytest tests/lanzadera/auth/test_issue_reset_token.py tests/lanzadera/auth/test_consume_reset_token.py --cov=app --cov-fail-under=85` | N/A | Revert use-case files | main |
| A3 | `feat(core): add admin use cases (create/disable user, grant/revoke admin, assign profile)` | `create_user`, `disable_user`, `grant_global_admin`, `revoke_global_admin`, `assign_profile`, `list_effective_apps`, `audit_append`, `bootstrap_global_admins`, `set_password` (CLI bootstrap); pure application layer | 260 | `pytest tests/lanzadera/application/` | N/A | Revert use-case files | main |
| AD1 | `feat(adapters): add Postgres session factory and user/app/profile repositories` | `async_session_factory` from asyncpg+SQLAlchemy 2.0 Core; `UserRepositoryPg` (dni_encrypted via NationalIdCipher), `AppRepositoryPg`, `ProfileRepositoryPg`; contract tests with in-memory fakes | 280 | `pytest tests/lanzadera/adapters/ --cov=app --cov-fail-under=85` | `docker compose up -d postgres` | Revert adapter files | main |
| AD2 | `feat(adapters): add assignment, reset_token, global_admin, audit repositories` | `AssignmentRepositoryPg`, `ResetTokenRepositoryPg`, `GlobalAdminRepositoryPg`, `AuditLogPg` (same-transaction atomic per DA-11); contract tests | 260 | `pytest tests/lanzadera/adapters/ --cov=app --cov-fail-under=85` | Same compose | Revert adapter files | main |
| AD3 | `feat(adapters): add Argon2id credential hasher and NationalId cipher` | `CredentialHasherArgon2id` implementing `CredentialHasherPort`; `NationalIdCipher` using AES-GCM via `SecretManagerPort`; key never in logs/CLI args (CA-S4); unit tests | 200 | `pytest tests/lanzadera/adapters/crypto/ --cov=app --cov-fail-under=85` | N/A | Revert crypto files | main |
| AD4 | `feat(adapters): add MailQueueTable and EnvBootstrap adapters` | `MailQueueTableAdapter` (DA-10, inserts mail_outbox pending); `EnvAdminSourceAdapter` reads `GLOBAL_ADMIN_EMAILS` (DA-6, idempotent); unit tests for both | 180 | `pytest tests/lanzadera/adapters/notification/ tests/lanzadera/test_bootstrap_adapter.py` | N/A | Revert adapter files | main |
| AD5 | `feat(adapters): add AssumeInOffice, TTLCache and EnvSecretManager adapters` | `AssumeInOfficeAdapter` (always True, DA-9/H12); `TtlCacheAdapter` using cachetools.TTLCache for apps/profiles/effective_perms; `EnvSecretManagerAdapter` (MVP); unit tests | 180 | `pytest tests/lanzadera/adapters/cache/ tests/lanzadera/test_assume_in_office.py` | N/A | Revert adapter files | main |
| AD6 | `feat(di): add composition root wiring and bootstrap` | `container.py`: wires all ports to adapters; FastAPI `Depends()` declarations; `bootstrap.py`: calls `bootstrap_global_admins` at startup + `purge_expired` in background; idempotent | 160 | `pytest tests/lanzadera/di/` | N/A | Revert DI files | main |
| DL1 | `feat(delivery): add HTTP router with login/logout and reset endpoints` | FastAPI router `/admin/`; `POST /login` (lockout after 5 failures D38), `POST /logout`, `GET /reset`, `POST /reset`; HTMX fragment responses; Alpine.js; RED tests for login/reset | 280 | `pytest tests/lanzadera/delivery/test_login_view.py tests/lanzadera/delivery/test_reset_view.py` | `docker compose up` | Remove delivery routes | main |
| DL2 | `feat(delivery): add admin users/apps/assignments/audit HTMX views` | `GET /admin/users`, `POST /admin/users`, `PATCH /admin/users/{id}/disable`, `POST /admin/assignments`, `GET /admin/apps`, `POST /admin/apps/{id}/activate`, `GET /admin/audit`; async Jinja templates with Mistica tokens | 300 | `pytest tests/lanzadera/delivery/` | Same compose | Remove delivery routes | main |
| DL3 | `feat(delivery): add platform user CLI with set-password/grant/revoke commands` | `gentle-ai platform user set-password <email>` (D91, admin-only, Argon2id), `grant-global-admin`, `revoke-global-admin`, `list-apps`, `assign-profile`; destructive commands prompt confirmation (D26); secrets via env (CA-S4); RED tests | 260 | `pytest tests/lanzadera/delivery/cli/` | N/A | Remove CLI module | main |
| T1 | `feat(tests): add unit and integration test suite` | `test_no_legacy_compat.py` (AST pin DA-13), `test_lockout_policy.py` (D38-D40), `test_same_transaction_audit.py` (DA-11 rollback), `test_secret_manager_contract.py`, `test_mail_queue_contract.py`, `test_grant_revoke_round_trip.py`, `test_gate_smoke.py` (Hard Rule 18); full suite run | 300 | `pytest --cov=app --cov-fail-under=85 -p app.pytest_plugin.coverage_gate` | Same compose | Revert test files | main |
| E1 | `test(e2e): add migration integration and login-reset E2E via httpx` | `alembic upgrade head` → all 6 migrations apply; httpx client: POST `/login` with seeded user → receives token via mail_outbox; POST `/reset` with token+password → user active; subsequent POST `/login` → 200 | 200 | `pytest tests/lanzadera/e2e/` | Full `docker compose up` | `docker compose down && docker compose up -d postgres` | main |
| E2 | `test(e2e): add CLI E2E and migrate_from_access script smoke` | CLI: `gentle-ai platform user set-password admin@empresa.com` → user active → issue_reset_token succeeds (no no_global_admin error); lockout after 5 failures in DB; `migrate_from_access.py` with SHA256 mismatch → aborts non-zero | 200 | `pytest tests/lanzadera/e2e/` | Full `docker compose up` | Revert CLI files | main |
| E3 | `test(e2e): add final quality gates verification` | No legacy_hash/sha256/verify_legacy/old_password/migrate_password in app/src/ (DA-13 green); `make quality-report` clean; `ruff format --check` passes; `mypy app/src/` passes; `ruff check app/` passes | 160 | `make quality-report` | Full `docker compose up` | `git checkout HEAD -- app/` | main |

## Phase 1: Foundation — Platform scaffold + CI (F1–F9)

### WU F1 — `feat(lanzadera): scaffold platform directory tree`

**Vertical slice**: Empty `app/` tree with `__init__.py` stubs for all future packages.
**Acceptance criteria**: All `__init__.py` present; `app/src/main.py` exists as stub; no `.py` files with real logic yet.
**Est. lines**: 80 · **Test**: N/A · **Harness**: N/A
**Rollback**: Delete `app/src/modules/lanzadera/`
**Base**: main

### WU F2 — `feat(ci): add pyproject.toml with exact dependency pins`

**Vertical slice**: `pyproject.toml` declaring exact pins (argon2-cffi==25.1.0, sqlalchemy==2.0.*, alembic==1.13.*, asyncpg==0.30.*, fastapi==0.119.*, jinja2==3.1.*, httpx==0.28.*, cachetools==5.*, ruff==0.15.21, mypy==1.13.0); ruff/mypy tool sections per QC-3/QC-4; pytest-cov config; `[tool.coverage]` with `app/` as source.
**Acceptance criteria**: `ruff check app/` and `mypy app/src/` run without error on the empty scaffold.
**Est. lines**: 120 · **Test**: `ruff check app/` · **Harness**: N/A
**Rollback**: Revert file
**Base**: main

### WU F3 — `feat(infra): add Dockerfile and docker-compose.yml`

**Vertical slice**: Multi-stage `Dockerfile` with `python:3.12-slim-bookworm@sha256:<digest>` (all FROMs pinned); `docker-compose.yml` with postgres + minio + backend service; `PLATFORM_SECRET_KEY` seeded in `secrets/` volume; `GLOBAL_ADMIN_EMAILS` env var documented.
**Acceptance criteria**: `docker compose up -d postgres` and `docker compose ps` show healthy postgres; backend service starts without error.
**Est. lines**: 140 · **Test**: `docker compose up -d postgres` · **Harness**: `docker compose ps`
**Rollback**: `docker compose down`
**Base**: main

### WU F4 — `feat(ci): add Makefile with all quality targets`

**Vertical slice**: `Makefile` with all targets: `lint`, `typecheck`, `test`, `check-layers`, `check-complexity`, `check-crap`, `check-dry`, `check-pr-size`, `check-branch-name`, `security`, `quality-report`, `migrate-fixtures`.
**Acceptance criteria**: `make lint` and `make typecheck` pass on empty scaffold; `make check-layers` returns 0 violations.
**Est. lines**: 100 · **Test**: `make lint` · **Harness**: N/A
**Rollback**: Revert file
**Base**: main

### WU F5 — `feat(ci): add GitHub Actions CI workflow`

**Vertical slice**: `.github/workflows/ci.yml` with all gates (ruff, mypy, pytest --cov, coverage_gate, check_layers, check_complexity, check_crap, check_dry, check_pr_size, check_branch_name, quality_report); `tests/test_ci_workflow.py` verifies every step is present.
**Acceptance criteria**: `pytest tests/test_ci_workflow.py` passes; all CI steps defined.
**Est. lines**: 160 · **Test**: `pytest tests/test_ci_workflow.py` · **Harness**: N/A
**Rollback**: `git revert HEAD`
**Base**: main

### WU F6 — `feat(security): add security scanning workflows`

**Vertical slice**: `.github/workflows/security.yml` with `pip-audit` (venv throwaway), `gitleaks` (digest pinned), `trivy config` (digest pinned); `.github/workflows/security-deep.yml` for weekly deep scan.
**Acceptance criteria**: `gitleaks detect --source=. --redact` returns 0 findings on the scaffold.
**Est. lines**: 140 · **Test**: `gitleaks detect --source=. --redact` · **Harness**: N/A
**Rollback**: `git revert HEAD`
**Base**: main

### WU F7 — `feat(qa): add hexagonal layer gate scripts`

**Vertical slice**: `scripts/check_layers.py` from `deterministic-quality-harness/assets/`; `ROOT_PACKAGE = "app.src.modules"`, `PURE_LAYERS = {"domain","ports","application"}`, `ALLOWED_IMPORTS` per DA-1; `tests/lanzadera/test_layers_wiring.py` verifying gate wiring.
**Acceptance criteria**: `python app/scripts/check_layers.py --root .` returns exit 0 on empty `lanzadera/`; `test_layers_wiring.py` passes.
**Est. lines**: 160 · **Test**: `python app/scripts/check_layers.py --root .` · **Harness**: N/A
**Rollback**: Revert script
**Base**: main

### WU F8 — `feat(qa): add complexity, CRAP, DRY and quality-report scripts`

**Vertical slice**: `check_complexity.py` (CC≤15 absolute ceiling per QC-10), `check_crap.py` (CRAP≤6), `check_dry.py` (0 clones), `quality_report.py` (layers→complexity→crap→dry fixed order); wiring tests for each.
**Acceptance criteria**: All scripts run on empty `lanzadera/` without error; `test_complexity_wiring.py` passes.
**Est. lines**: 200 · **Test**: `python app/scripts/check_complexity.py --root .` · **Harness**: N/A
**Rollback**: Revert scripts
**Base**: main

### WU F9 — `feat(qa): add PR-size and branch-name gate scripts`

**Vertical slice**: `check_pr_size.py` (400-line gate), `check_branch_name.py` (regex `^(feat|fix|refactor|docs|ci|test)/<n>-<slug>$`), `coverage_gate.py` declaring `CRITICAL_HELPERS = ["hash_password","verify_password","issue_reset_token","consume_reset_token"]`; smoke test via `test_gate_smoke.py`.
**Acceptance criteria**: Each script exits 0 on the scaffold; `test_gate_smoke.py` passes.
**Est. lines**: 120 · **Test**: `python app/scripts/check_pr_size.py --root .` · **Harness**: N/A
**Rollback**: Revert scripts
**Base**: main

## Phase 2: Domain Layer — Pure entities (D1–D4)

### WU D1 — `feat(domain): add User, Session, ResetToken entities`

**Vertical slice**: Pure domain classes with no framework imports. `User` entity with all fields (email normalised lowercase, non-empty name invariants), `UserStatus` StrEnum, `Session` entity, `ResetToken` VO.
**Acceptance criteria**: `check_layers.py` passes on `domain/` (PURE_LAYERS); all entities pass their own unit tests.
**Est. lines**: 160 · **Test**: `pytest tests/lanzadera/domain/` · **Harness**: N/A
**Rollback**: Delete domain files
**Base**: main

### WU D2 — `feat(domain): add App, Profile, Assignment entities`

**Vertical slice**: Pure domain. `App` with `AppTopology`/`AppRegistrationStatus` ENUMs, `Profile` with `capabilities` JSONB shape, `Assignment` with `granted_by` nullable.
**Acceptance criteria**: `check_layers.py` passes; unit tests for entity invariants.
**Est. lines**: 140 · **Test**: `pytest tests/lanzadera/domain/` · **Harness**: N/A
**Rollback**: Delete domain files
**Base**: main

### WU D3 — `feat(domain): add GlobalAdmin, AuditEvent, legacy_role_map`

**Vertical slice**: Pure domain. `GlobalAdmin` VO; `AuditEvent` (no telemetry per DA-11); `legacy_role_map.py` covering 7 legacy flags → profile codes, SinAcceso cortocircuito, 32-case matrix (H11). G-2 ABIERTO: seed with `capabilities={}`.
**Acceptance criteria**: `test_sinacceso_exclusivity.py` covers all 32 cases; `check_layers.py` passes.
**Est. lines**: 180 · **Test**: `pytest tests/lanzadera/domain/` · **Harness**: N/A
**Rollback**: Delete domain files
**Base**: main

### WU D4 — `feat(domain): add LockoutPolicy value object`

**Vertical slice**: Pure domain. `LockoutPolicy(threshold: int = 5, duration_seconds: int = 3600)`.
**Acceptance criteria**: `check_layers.py` passes; unit tests.
**Est. lines**: 80 · **Test**: `pytest tests/lanzadera/domain/` · **Harness**: N/A
**Rollback**: Delete domain files
**Base**: main

## Phase 3: Migrations (M1–M6)

### WU M1 — `chore(migrations): add 0001_core_schema`

**Vertical slice**: Creates `lanzadera` schema; all 8 tables; ENUMs; unique constraints; indexes; **no** telemetry columns (DA-11).
**Acceptance criteria**: `alembic upgrade head` applies without error; `alembic current` shows 0001.
**Est. lines**: 220 · **Test**: `alembic upgrade head` · **Harness**: Fresh `initdb`
**Rollback**: `alembic downgrade -1`
**Base**: main

### WU M2 — `chore(migrations): add 0002_seed_apps`

**Vertical slice**: Seed 20 apps from fixture; excludes Pass/Comando/URLDIrectorioIconoAplicacion (DA-7, closes G-3). RED test first.
**Acceptance criteria**: `pytest tests/lanzadera/migrations/test_migration_0002.py` passes; exactly 20 rows in `apps`; no launcher fields.
**Est. lines**: 140 · **Test**: `alembic upgrade head && pytest tests/lanzadera/migrations/` · **Harness**: Same compose
**Rollback**: `alembic downgrade -1`
**Base**: main

### WU M3 — `chore(migrations): add 0003_seed_profiles`

**Vertical slice**: Seed `default` profile per app + legacy-role profiles per app. `capabilities={}` (G-2 ABIERTO). RED test first.
**Acceptance criteria**: Each app has at least one profile with `code='default'`; pytest passes.
**Est. lines**: 120 · **Test**: `alembic upgrade head && pytest tests/lanzadera/migrations/` · **Harness**: Same compose
**Rollback**: `alembic downgrade -1`
**Base**: main

### WU M4 — `chore(migrations): add 0004_seed_users`

**Vertical slice**: Migrate 156 users; `email=lower(input.email)`; `dni_encrypted=encrypt(input.dni)`; `password_hash=NULL`; `status='password_reset_required'`; NO legacy_hash (DA-3). RED test first.
**Acceptance criteria**: `password_hash IS NULL` and `status='password_reset_required'` for all migrated users; no `legacy_hash` column.
**Est. lines**: 160 · **Test**: `alembic upgrade head && pytest tests/lanzadera/migrations/` · **Harness**: Same compose
**Rollback**: `alembic downgrade -1`
**Base**: main

### WU M5 — `chore(migrations): add 0005_seed_assignments`

**Vertical slice**: Apply legacy-role mapping; SinAcceso exclusive cortocircuito; DEFAULT for all-NULL; ≈622 rows (H11). RED test first.
**Acceptance criteria**: `test_sinacceso_exclusivity.py` passes all 32 cases; cardinality ≈622.
**Est. lines**: 140 · **Test**: `alembic upgrade head && pytest tests/lanzadera/assignments/` · **Harness**: Same compose
**Rollback**: `alembic downgrade -1`
**Base**: main

### WU M6 — `chore(migrations): add 0006_seed_audit`

**Vertical slice**: Seed TbConexiones → auth.login.success/failure; seed TbAplicacionesAperturas → app.open; NO SSID/BS/coords/machine/IP (D55). RED test first.
**Acceptance criteria**: Audit rows present; no telemetry columns populated.
**Est. lines**: 120 · **Test**: `alembic upgrade head && pytest tests/lanzadera/audit/` · **Harness**: Same compose
**Rollback**: `alembic downgrade -1`
**Base**: main

## Phase 4: Application Layer — Use cases + ports (A1–A3)

### WU A1 — `feat(auth): add hash_password and verify_password (CRITICAL_HELPERS)`

**Vertical slice**: RED first. `hash_password` produces Argon2id PHC string (m=65536, t=3, p=4); two calls produce different hashes; `verify_password` returns True/False; raises `ValueError` on empty/nil hash. 100% coverage gate.
**Acceptance criteria**: `pytest tests/lanzadera/auth/test_hash_password.py tests/lanzadera/auth/test_verify_password.py --cov=app --cov-fail-under=85` passes; CRITICAL_HELPERS at 100%.
**Est. lines**: 180 · **Test**: `pytest tests/lanzadera/auth/test_hash_password.py tests/lanzadera/auth/test_verify_password.py --cov=app --cov-fail-under=85` · **Harness**: N/A
**Rollback**: Revert helper files
**Base**: main

### WU A2 — `feat(auth): add issue_reset_token and consume_reset_token (CRITICAL_HELPERS)`

**Vertical slice**: RED first. `issue_reset_token`: 24h expiry, supersedes prior, NoGlobalAdminError guard, calls NotificationDeliveryPort.send once. `consume_reset_token`: atomic, single-use, rollback on audit failure. 100% coverage gate.
**Acceptance criteria**: `pytest tests/lanzadera/auth/test_issue_reset_token.py tests/lanzadera/auth/test_consume_reset_token.py --cov=app --cov-fail-under=85` passes; atomic rollback test green.
**Est. lines**: 280 · **Test**: `pytest tests/lanzadera/auth/test_issue_reset_token.py tests/lanzadera/auth/test_consume_reset_token.py --cov=app --cov-fail-under=85` · **Harness**: N/A
**Rollback**: Revert use-case files
**Base**: main

### WU A3 — `feat(core): add admin use cases`

**Vertical slice**: `create_user`, `disable_user`, `grant_global_admin`, `revoke_global_admin`, `assign_profile`, `list_effective_apps`, `audit_append`, `bootstrap_global_admins`, `set_password` (CLI bootstrap). Pure application layer, no adapters imported.
**Acceptance criteria**: `check_layers.py` passes (no adapter imports in application layer); unit tests for each use case.
**Est. lines**: 260 · **Test**: `pytest tests/lanzadera/application/` · **Harness**: N/A
**Rollback**: Revert use-case files
**Base**: main

## Phase 5: Adapters — Driven implementations (AD1–AD6)

### WU AD1 — `feat(adapters): add Postgres session factory and user/app/profile repositories`

**Vertical slice**: `async_session_factory` from asyncpg+SQLAlchemy 2.0 Core bound to `DATABASE_URL`; `UserRepositoryPg` (dni_encrypted via NationalIdCipher before INSERT); `AppRepositoryPg`; `ProfileRepositoryPg`; in-memory fakes for contract tests.
**Acceptance criteria**: `pytest tests/lanzadera/adapters/ --cov=app --cov-fail-under=85` passes; contract tests green.
**Est. lines**: 280 · **Test**: `pytest tests/lanzadera/adapters/ --cov=app --cov-fail-under=85` · **Harness**: `docker compose up -d postgres`
**Rollback**: Revert adapter files
**Base**: main

### WU AD2 — `feat(adapters): add assignment, reset_token, global_admin, audit repositories`

**Vertical slice**: `AssignmentRepositoryPg`; `ResetTokenRepositoryPg` (index on user_id+expires_at); `GlobalAdminRepositoryPg` (revoke rejects last admin constraint); `AuditLogPg` (same-transaction atomic per DA-11); contract tests.
**Acceptance criteria**: `pytest tests/lanzadera/adapters/ --cov=app --cov-fail-under=85` passes; `test_same_transaction_audit.py` green.
**Est. lines**: 260 · **Test**: `pytest tests/lanzadera/adapters/ --cov=app --cov-fail-under=85` · **Harness**: Same compose
**Rollback**: Revert adapter files
**Base**: main

### WU AD3 — `feat(adapters): add Argon2id credential hasher and NationalId cipher`

**Vertical slice**: `CredentialHasherArgon2id` implementing `CredentialHasherPort`; `NationalIdCipher` AES-GCM via `SecretManagerPort`; key never in logs/CLI args (CA-S4). Tests delegate to A1 helpers.
**Acceptance criteria**: `pytest tests/lanzadera/adapters/crypto/ --cov=app --cov-fail-under=85` passes.
**Est. lines**: 200 · **Test**: `pytest tests/lanzadera/adapters/crypto/ --cov=app --cov-fail-under=85` · **Harness**: N/A
**Rollback**: Revert crypto files
**Base**: main

### WU AD4 — `feat(adapters): add MailQueueTable and EnvBootstrap adapters`

**Vertical slice**: `MailQueueTableAdapter` (DA-10, inserts mail_outbox with `status='pending'`); `EnvAdminSourceAdapter` reads `GLOBAL_ADMIN_EMAILS` (DA-6, idempotent, empty if unset); unit tests for both.
**Acceptance criteria**: `pytest tests/lanzadera/adapters/notification/ tests/lanzadera/test_bootstrap_adapter.py` passes.
**Est. lines**: 180 · **Test**: `pytest tests/lanzadera/adapters/notification/ tests/lanzadera/test_bootstrap_adapter.py` · **Harness**: N/A
**Rollback**: Revert adapter files
**Base**: main

### WU AD5 — `feat(adapters): add AssumeInOffice, TTLCache and EnvSecretManager adapters`

**Vertical slice**: `AssumeInOfficeAdapter` (always True, DA-9/H12); `TtlCacheAdapter` using cachetools.TTLCache for apps.list_active, profiles.list_for_app, assignments.effective_permissions; `EnvSecretManagerAdapter` (MVP); unit tests.
**Acceptance criteria**: `pytest tests/lanzadera/adapters/cache/ tests/lanzadera/test_assume_in_office.py` passes.
**Est. lines**: 180 · **Test**: `pytest tests/lanzadera/adapters/cache/ tests/lanzadera/test_assume_in_office.py` · **Harness**: N/A
**Rollback**: Revert adapter files
**Base**: main

### WU AD6 — `feat(di): add composition root wiring and bootstrap`

**Vertical slice**: `container.py`: wires all ports to adapters; FastAPI `Depends()` declarations; `bootstrap.py`: calls `bootstrap_global_admins` at startup + `purge_expired` in background; idempotent (DA-6).
**Acceptance criteria**: `check_layers.py` passes; `pytest tests/lanzadera/di/` passes.
**Est. lines**: 160 · **Test**: `pytest tests/lanzadera/di/` · **Harness**: N/A
**Rollback**: Revert DI files
**Base**: main

## Phase 6: Delivery — HTTP HTMX + CLI (DL1–DL3)

### WU DL1 — `feat(delivery): add HTTP router with login/logout and reset endpoints`

**Vertical slice**: FastAPI router `/admin/`; `POST /login` (lockout after 5 failures D38), `POST /logout`, `GET /reset`, `POST /reset`; HTMX fragment responses with `HX-Retarget`; Alpine.js. RED tests for login/reset first.
**Acceptance criteria**: `pytest tests/lanzadera/delivery/test_login_view.py tests/lanzadera/delivery/test_reset_view.py` passes; POST `/login` with wrong password returns 401; 5 bad attempts locks account.
**Est. lines**: 280 · **Test**: `pytest tests/lanzadera/delivery/test_login_view.py tests/lanzadera/delivery/test_reset_view.py` · **Harness**: `docker compose up`
**Rollback**: Remove delivery routes
**Base**: main

### WU DL2 — `feat(delivery): add admin users/apps/assignments/audit HTMX views`

**Vertical slice**: `GET /admin/users`, `POST /admin/users`, `PATCH /admin/users/{id}/disable`, `POST /admin/assignments`, `GET /admin/apps`, `POST /admin/apps/{id}/activate`, `GET /admin/audit`; async Jinja templates with Mistica token CSS vars.
**Acceptance criteria**: `pytest tests/lanzadera/delivery/` passes; all routes return valid HTMX fragments.
**Est. lines**: 300 · **Test**: `pytest tests/lanzadera/delivery/` · **Harness**: Same compose
**Rollback**: Remove delivery routes
**Base**: main

### WU DL3 — `feat(delivery): add platform user CLI with set-password/grant/revoke commands`

**Vertical slice**: `gentle-ai platform user set-password <email>` (D91, admin-only, Argon2id), `grant-global-admin`, `revoke-global-admin`, `list-apps`, `assign-profile`; destructive commands prompt confirmation (D26); secrets via env (CA-S4). RED tests first.
**Acceptance criteria**: `pytest tests/lanzadera/delivery/cli/` passes; `set-password` creates user with Argon2id hash, sets `status='active'`, adds to global_admins; idempotent on second run; rejects non-admin caller.
**Est. lines**: 260 · **Test**: `pytest tests/lanzadera/delivery/cli/` · **Harness**: N/A
**Rollback**: Remove CLI module
**Base**: main

## Phase 7: Tests + Gates (T1)

### WU T1 — `feat(tests): add unit and integration test suite`

**Vertical slice**: `test_no_legacy_compat.py` (AST pin DA-13, RED first — add dummy `legacy_hash = ...` and assert failure); `test_lockout_policy.py` (D38-D40); `test_same_transaction_audit.py` (DA-11 rollback); `test_secret_manager_contract.py`; `test_mail_queue_contract.py`; `test_grant_revoke_round_trip.py`; `test_gate_smoke.py` (Hard Rule 18 degenerate-run guard). Full suite run.
**Acceptance criteria**: `pytest --cov=app --cov-fail-under=85 -p app.pytest_plugin.coverage_gate` passes; global ≥85%, CRITICAL_HELPERS at 100%.
**Est. lines**: 300 · **Test**: `pytest --cov=app --cov-fail-under=85 -p app.pytest_plugin.coverage_gate` · **Harness**: `docker compose up`
**Rollback**: Revert test files
**Base**: main

## Phase 8: E2E Verification (E1–E3)

### WU E1 — `test(e2e): add migration integration and login-reset E2E via httpx`

**Vertical slice**: `alembic upgrade head` applies all 6 migrations; httpx test client: POST `/login` with seeded user (status=`password_reset_required`) → receives token via mail_outbox; POST `/reset` with token + new password → user becomes `active`; subsequent POST `/login` with new password → 200.
**Acceptance criteria**: All E2E steps pass; `make quality-report` clean.
**Est. lines**: 200 · **Test**: `pytest tests/lanzadera/e2e/` · **Harness**: Full `docker compose up`
**Rollback**: `docker compose down && docker compose up -d postgres`
**Base**: main

### WU E2 — `test(e2e): add CLI E2E and migrate_from_access script smoke`

**Vertical slice**: CLI: `gentle-ai platform user set-password admin@empresa.com` → user active → issue_reset_token succeeds (no no_global_admin error); lockout after 5 failures verified in DB; `migrate_from_access.py` with SHA256 mismatch → aborts non-zero.
**Acceptance criteria**: CLI E2E steps pass; SHA256 mismatch test aborts correctly.
**Est. lines**: 200 · **Test**: `pytest tests/lanzadera/e2e/` · **Harness**: Full `docker compose up`
**Rollback**: Revert CLI files
**Base**: main

### WU E3 — `test(e2e): add final quality gates verification`

**Vertical slice**: No `legacy_hash`/`sha256`/`verify_legacy`/`old_password`/`migrate_password` in `app/src/` (DA-13 AST test green); `make quality-report` clean; `ruff format --check` passes; `mypy app/src/` passes with `enable_error_code = ["ignore-without-code"]`; `ruff check app/` passes.
**Acceptance criteria**: All commands exit 0; quality gates evidence bundle published.
**Est. lines**: 160 · **Test**: `make quality-report` · **Harness**: Full `docker compose up`
**Rollback**: `git checkout HEAD -- app/`
**Base**: main

## Gap Tracking

| Gap | Status | Owner |
|-----|--------|-------|
| G-1 (email normalisation) | **Closed** in DA-3 — `lower(input.email)` | — |
| G-2 (capability set per app) | **Open — product must decide** before 0003 seed; seed `capabilities = {}` for now | Product |
| G-3 (apps fields not migrated) | **Closed** in DA-7 / M2 — `Pass`, `Comando`, `URLDIrectorioIconoAplicacion` excluded | — |
| G-4 (audit policy for failed admin creation) | **Open** — log at `WARN` severity in MVP; SOC notification deferred | SOC team |
| G-5 (SMTP notification channel) | **Open** — `MailQueueTableAdapter` MVP; SMTP real at P20 | IT |
| G-6 (password expiry policy) | **Open** — not implemented in MVP; `password_reset_required` covers recovery | Product |
| G-7 (SinAcceso weight in menu) | **Open** — stub menu filter returns all visible apps; behaviour decided in UI epic | UI epic |
| G-8 (audit retention periods) | **Open** — MVP: 90 days hot + 1 year total; `AuditRetentionPort` stub | Compliance |

> **Note**: G-2, G-4, G-5, G-6, G-7, G-8 are **non-blocking** for the apply phase. Each is either addressed by a stub/placeholder or deferred to a later release with the decision owner noted.

## Chain Strategy

**Recommended: `stacked-to-main`**

Each WU is a self-contained vertical slice that lands independently to `main`. No WU depends on a previous WU being in `main` before it can be reviewed and merged — the WU's own diff is the full scope, tests travel with code, and CI proves the slice works in isolation.

The tracker branch (`feat/lanzadera-mvp-tracker`) is used only as a convenient holding area for the pre-merge stack visualization. Only the WU branches (e.g. `feat/0001-F1-scaffold`) merge to `main`; the tracker never does.

```text
main
  └─► feat/0001-F1-scaffold          ──► PR → main
        └─► feat/0002-F2-pyproject      ──► PR → main
              └─► feat/0003-F3-docker      ──► PR → main
                    └─► feat/0004-F4-makefile  ──► PR → main
                          └─► feat/0005-F5-ci-workflow  ──► PR → main
                                └─► feat/0006-F6-security-wf  ──► PR → main
                                      └─► feat/0007-F7-layer-gate  ──► PR → main
                                            └─► feat/0008-F8-complexity-gates ──► PR → main
                                                  └─► feat/0009-F9-pr-size-gate ──► PR → main
                                                        └─► feat/0010-D1-user-domain ──► PR → main
```

(Pattern continues: D2→D3→D4→M1→M2→M3→M4→M5→M6→C1→A1→A2→A3→AD1→AD2→AD3→AD4→AD5→AD6→DL1→DL2→DL3→T1→E1→E2→E3 → main)

**Why not Feature Branch Chain?** Every WU is independently valuable and reviewable. Migrations M1–M6 are additive and don't require prior WUs to be in the tracker. Using stacked-to-main keeps the history clean and avoids a "waterfall review" bottleneck at the tracker branch.

**Why not a single PR?** Original PR 1 was 4224 lines and PR 8 was 400+ lines — both rejected. Every WU now fits ≤320 lines. No size:exception needed.

## Delivery Units — WU Mapping

| WU | Phase | Focus | Est. lines | Base |
|----|-------|-------|------------|------|
| F1 | Foundation | Platform scaffold directory tree | 80 | main |
| F2 | Foundation | pyproject.toml + pins | 120 | main |
| F3 | Foundation | Dockerfile + docker-compose.yml | 140 | main |
| F4 | Foundation | Makefile all quality targets | 100 | main |
| F5 | Foundation | CI workflow + gate wiring test | 160 | main |
| F6 | Foundation | Security scanning workflows | 140 | main |
| F7 | Foundation | Hexagonal layer gate scripts | 160 | main |
| F8 | Foundation | Complexity/CRAP/DRY/quality-report scripts | 200 | main |
| F9 | Foundation | PR-size/branch-name/coverage_gate scripts | 120 | main |
| D1 | Domain | User, Session, ResetToken entities | 160 | main |
| D2 | Domain | App, Profile, Assignment entities | 140 | main |
| D3 | Domain | GlobalAdmin, AuditEvent, legacy_role_map | 180 | main |
| D4 | Domain | LockoutPolicy value object | 80 | main |
| M1 | Migrations | 0001_core_schema | 220 | main |
| M2 | Migrations | 0002_seed_apps | 140 | main |
| M3 | Migrations | 0003_seed_profiles | 120 | main |
| M4 | Migrations | 0004_seed_users | 160 | main |
| M5 | Migrations | 0005_seed_assignments | 140 | main |
| M6 | Migrations | 0006_seed_audit | 120 | main |
| C1 | Scripts | migrate_from_access.py one-shot | 280 | main |
| A1 | Application | hash_password + verify_password | 180 | main |
| A2 | Application | issue_reset_token + consume_reset_token | 280 | main |
| A3 | Application | Admin use cases | 260 | main |
| AD1 | Adapters | Postgres session + user/app/profile repos | 280 | main |
| AD2 | Adapters | Assignment/reset_token/global_admin/audit repos | 260 | main |
| AD3 | Adapters | Argon2id hasher + NationalId cipher | 200 | main |
| AD4 | Adapters | MailQueueTable + EnvBootstrap adapters | 180 | main |
| AD5 | Adapters | AssumeInOffice + TTLCache + EnvSecretManager | 180 | main |
| AD6 | DI | Composition root + bootstrap | 160 | main |
| DL1 | Delivery | HTTP router login/logout/reset | 280 | main |
| DL2 | Delivery | Admin HTMX views | 300 | main |
| DL3 | Delivery | Platform user CLI | 260 | main |
| T1 | Tests | Unit/integration suite + gate smoke | 300 | main |
| E1 | E2E | Migration integration + login-reset E2E | 200 | main |
| E2 | E2E | CLI E2E + migrate script smoke | 200 | main |
| E3 | E2E | Final quality gates verification | 160 | main |

> **Total: 36 WUs. Max lines: 320. No size:exception.**
