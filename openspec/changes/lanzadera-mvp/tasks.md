# Tasks: Lanzadera MVP — Primer slice de la plataforma hexagonal

> Change: `lanzadera-mvp` · Project: `access2web-blueprint` · strict_tdd: true · pytest: `pytest --cov=app --cov-fail-under=85`

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 2 800–3 200 (auth alone ~600; migrations ~700; adapters ~800; delivery ~500; tests ~600) |
| 400-line budget risk | **High** |
| Chained PRs recommended | **Yes** |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: pending
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Focused test command | Runtime harness | Rollback boundary |
|------|------|-----------|----------------------|-----------------|-------------------|
| 1 | Phase 0 — platform scaffold + CI gates | PR 1 (feat/0001-platform-skeleton) | `pytest --cov=app --cov-fail-under=85` | `docker compose up -d postgres` | Drop schema + remove `app/` directory |
| 2 | Phase 1 — domain entities | PR 2 (feat/0002-domain-entities) | `pytest tests/lanzadera/` | Same compose | Delete domain entities |
| 3 | Phase 2 — migrations 0001–0006 | PR 3 (feat/0003-alembic-migrations) | `alembic upgrade head` | Fresh `initdb` | `alembic downgrade -1` per migration |
| 4 | Phase 3 — application layer (CRITICAL_HELPERS) | PR 4 (feat/0004-application-use-cases) | `pytest --cov=app --cov-fail-under=85` | Same compose | Revert to stub use cases |
| 5 | Phase 5 — reset flow + lockout (D90, D38–D40) | PR 5 (feat/0005-reset-lockout) | `pytest tests/lanzadera/auth/` | Same compose | Rollback token table rows |
| 6 | Phase 4 — adapters (persistence, crypto, notification, cache, location, bootstrap) | PR 6 (feat/0006-adapters) | `pytest tests/lanzadera/` | Same compose | Delete adapter files |
| 7 | Phase 5 — delivery HTMX + CLI | PR 7 (feat/0007-delivery-cli) | `pytest tests/lanzadera/delivery/` | `docker compose up` | Remove delivery routes |
| 8 | Phase 6 — all tests + gates verification | PR 8 (feat/0008-tests-gates) | `make quality-report` | Full E2E via httpx test client | Re-run test suite |

## Phase 0: Foundation — Platform scaffold + CI gates (TK-LZ-MVP-1, TK-LZ-MVP-2, TK-LZ-MVP-3)

- [ ] 0.1 Create `app/` directory tree: `src/modules/lanzadera/{domain,ports,application,adapters,di,delivery}`, `migrations/versions/`, `tests/lanzadera/`, `scripts/`, `src/shared/`
- [ ] 0.2 Create `app/pyproject.toml` with exact pins: `argon2-cffi==25.1.0`, `sqlalchemy==2.0.*`, `alembic==1.13.*`, `asyncpg==0.30.*`, `fastapi==0.119.*`, `jinja2==3.1.*`, `httpx==0.28.*`, `cachetools==5.*`, `ruff==0.15.21`, `mypy==1.13.0`; `pytest-cov`; `[tool.coverage]` with `app/` as source
- [ ] 0.3 Create `app/Dockerfile` with `python:3.12-slim-bookworm@sha256:<digest>`; multi-stage build; pin every `FROM` digest
- [ ] 0.4 Create `docker-compose.yml` with postgres + minio + backend service; `PLATFORM_SECRET_KEY` seeded in `secrets/` volume; `GLOBAL_ADMIN_EMAILS` env var documented
- [ ] 0.5 Create `Makefile` with targets: `lint`, `typecheck`, `test`, `check-layers`, `check-complexity`, `check-crap`, `check-dry`, `check-pr-size`, `check-branch-name`, `security`, `quality-report`, `migrate-fixtures`
- [ ] 0.6 Create `.github/workflows/ci.yml` with all gates: `ruff check`, `mypy`, `pytest --cov=app --cov-fail-under=85`, `python scripts/check_layers.py`, `python scripts/check_complexity.py`, `python scripts/check_crap.py`, `python scripts/check_dry.py`, `python scripts/check_pr_size.py`, `python scripts/check_branch_name.py`, `python scripts/quality_report.py`
- [ ] 0.7 Create `.github/workflows/security.yml` with `pip-audit` (venv throwaway), `gitleaks` (digest pinned), `trivy config` (digest pinned); `.github/workflows/security-deep.yml` for weekly deep scan
- [ ] 0.8 Wire `tests/test_ci_workflow.py` to assert every CI step is present (pin test for gate wiring per QC-9 / Hard Rule 4 of deterministic-quality-harness)
- [ ] 0.9 Create `scripts/check_layers.py` from `deterministic-quality-harness/assets/`; set `ROOT_PACKAGE = "app.src.modules"`, `PURE_LAYERS = {"domain","ports","application"}`, `ALLOWED_IMPORTS` per DA-1; create `tests/lanzadera/test_layers_wiring.py`
- [ ] 0.10 Create `scripts/check_complexity.py` with `CC ≤ 15` absolute ceiling per QC-10; `tests/lanzadera/test_complexity_wiring.py`
- [ ] 0.11 Create `scripts/check_crap.py` with `CRAP ≤ 6`; `scripts/check_dry.py` with 0 tolerated clones; `scripts/quality_report.py` executing gates in fixed order
- [ ] 0.12 Create `scripts/check_pr_size.py` and `scripts/check_branch_name.py` per QC-6; `app/pytest_plugin/coverage_gate.py` declaring `CRITICAL_HELPERS = ["hash_password","verify_password","issue_reset_token","consume_reset_token"]`
- [ ] 0.13 Create `app/.python-version` = `3.12`; `.gitignore`, `.dockerignore` covering `__pycache__`, `.ruff_cache`, `.mypy_cache`, `.pytest_cache`, `.venv*`, `secrets/`
- [ ] 0.14 Create `app/src/__init__.py`, `app/src/shared/__init__.py`, `app/src/modules/__init__.py`, `app/src/modules/lanzadera/__init__.py`, `app/src/main.py` (FastAPI app stub with `get_current_admin` dependency stub)

## Phase 1: Domain Layer (pure, framework-free)

- [ ] 1.1 Create `app/src/modules/lanzadera/domain/user.py` with `User` entity: `id: UUID`, `email: str`, `name: str`, `dni_encrypted: bytes`, `password_hash: str | None`, `status: UserStatus` (ENUM `active,disabled,password_reset_required,locked`), `failed_attempts: int`, `last_login_at: datetime | None`, `created_at`, `updated_at`; `UserStatus` as `StrEnum`; domain invariants: email normalised lowercase, non-empty name
- [ ] 1.2 Create `app/src/modules/lanzadera/domain/app.py` with `App` entity: `id: int`, `name: str`, `short_code: str`, `deployment_topology: AppTopology` (ENUM `central,office-nas`), `requires_office_presence: bool`, `registration_status: AppRegistrationStatus` (ENUM `pending,active,retired`), `created_at`, `updated_at`
- [ ] 1.3 Create `app/src/modules/lanzadera/domain/profile.py` with `Profile` entity: `id: UUID`, `app_id: int`, `code: str`, `name: str`, `capabilities: dict` (JSONB shape `{string: string|number|boolean}`), `active: bool`, `created_at`, `updated_at`
- [ ] 1.4 Create `app/src/modules/lanzadera/domain/assignment.py` with `Assignment` entity: `id: UUID`, `user_id: UUID`, `app_id: int`, `profile_id: UUID`, `granted_by: UUID | None`, `granted_at: datetime`, `revoked_at: datetime | None`
- [ ] 1.5 Create `app/src/modules/lanzadera/domain/reset_token.py` with `ResetToken` value object: `id: UUID`, `user_id: UUID`, `token_hash: str`, `expires_at: datetime`, `consumed_at: datetime | None`, `superseded_at: datetime | None`, `created_at`
- [ ] 1.6 Create `app/src/modules/lanzadera/domain/global_admin.py` with `GlobalAdmin` value object: `user_id: UUID`
- [ ] 1.7 Create `app/src/modules/lanzadera/domain/audit_event.py` with `AuditEvent` entity: `id: UUID`, `event_type: str`, `actor_id: UUID | None`, `target_id: str`, `module: str`, `result: str`, `correlation_id: UUID`, `payload: dict`, `created_at`; **no** telemetry fields (DA-11)
- [ ] 1.8 Create `app/src/modules/lanzadera/domain/legacy_role_map.py` (DA-12, closes G-2): map of 7 legacy flags (`Administrador`, `Calidad`, `CalidadAvisos`, `Técnico`, `Economía`, `Secretaría`, `SinAcceso`) to profile codes; `SinAcceso` exclusive cortocircuito; `DEFAULT` for all-NULL; unit-tested against 32-case matrix (H11); **Note**: G-2 capability set per app is ABIERTO — seed with `capabilities = {}` or minimal shape; product must confirm canonical set before 0003
- [ ] 1.9 Create `app/src/modules/lanzadera/domain/session.py` with `Session` entity: `id: UUID`, `user_id: UUID`, `created_at`, `expires_at`; `LockoutPolicy` value object: `threshold: int` (default 5), `duration_seconds: int` (default 3600)

## Phase 2: Persistence — Alembic migrations 0001–0006 (TK-LZ-MVP-8 to TK-LZ-MVP-13)

- [ ] 2.1 **RED**: write `tests/lanzadera/migrations/test_migration_0001.py` asserting `users` table has no `legacy_hash` column (DA-3); asserting ENUMs `user_status`, `app_topology`, `app_registration_status` exist; asserting `reset_tokens` and `mail_outbox` tables exist
- [ ] 2.2 Create `migrations/versions/0001_core_schema.py`: creates `lanzadera` schema; tables `users`, `apps`, `profiles`, `user_app_assignments`, `global_admins`, `sessions`, `audit`, `reset_tokens`, `mail_outbox`; ENUMs; unique constraints; indexes; **no** telemetry columns (DA-11); rollback drops all tables
- [ ] 2.3 **RED**: write `tests/lanzadera/migrations/test_migration_0002.py` asserting exactly 20 rows in `apps` after seed; no launcher fields
- [ ] 2.4 Create `migrations/versions/0002_seed_apps.py`: seed 20 apps from `TbAplicaciones.json` fixture; maps `EjecucionEnOficina` → `deployment_topology`/`requires_office_presence`; **excludes** `Pass`, `Comando`, `URLDIrectorioIconoAplicacion` (DA-7, closes G-3)
- [ ] 2.5 **RED**: write `tests/lanzadera/migrations/test_migration_0003.py` asserting each app has at least one `profile` with `code = 'default'`
- [ ] 2.6 Create `migrations/versions/0003_seed_profiles.py`: seeds `default` profile per app; seeds legacy-role profiles (`ADMIN`, `CALIDAD`, `CALIDAD_AVISOS`, `TECNICO`, `ECONOMIA`, `SECRETARIA`, `SIN_ACCESO`) per app when legacy had those flags; `capabilities = {}` or minimal shape (G-2 ABIERTO — product must confirm)
- [ ] 2.7 **RED**: write `tests/lanzadera/migrations/test_migration_0004.py` asserting `password_hash IS NULL` and `status = 'password_reset_required'` for all migrated users; asserting no `legacy_hash` column
- [ ] 2.8 Create `migrations/versions/0004_seed_users.py`: migrate 156 users with `email = lower(input.email)`, `dni_encrypted = encrypt(input.dni)` via `NationalIdCipher`; `password_hash = NULL`, `status = 'password_reset_required'` (D89, DA-3); **no** `legacy_hash` column
- [ ] 2.9 **RED**: write `tests/lanzadera/assignments/test_sinacceso_exclusivity.py` — covers all 32 legacy combinations; asserts `SinAcceso` cortocircuito; asserts `DEFAULT` for all-NULL; asserts cardinality ≈ 622 rows (H11, DA-12)
- [ ] 2.10 Create `migrations/versions/0005_seed_assignments.py`: apply legacy-role mapping from `legacy_role_map.py`; `SinAcceso` exclusive rule; compound rows; `DEFAULT` fallback
- [ ] 2.11 Create `migrations/versions/0006_seed_audit.py`: seed `TbConexiones` → `auth.login.success`/`auth.login.failure`; seed `TbAplicacionesAperturas` → `app.open`; **no** SSID/BS coordinates/machine/IP (D55); use fixture files with SHA256 verified by `migrate_from_access.py`
- [ ] 2.12 Create `scripts/migrate_from_access.py` (one-shot): Dysflow read-only extract → JSON fixtures; verifies SHA256 of `.accdb` before processing; produces `TbAplicaciones.json`, `tbUsuarios.json`, `TbUsuariosAplicacionesPermisos.json`, `TbConexiones.json`, `TbAplicacionesAperturas.json`; callable via `make migrate-fixtures`

## Phase 3: Application Layer — Use cases + ports

- [ ] 3.1 Create `app/src/modules/lanzadera/ports/user_repository.py` Protocol: `get_by_email`, `create`, `update_status`, `set_password_and_activate`, `list_all`, `get_by_id`; `app/src/modules/lanzadera/ports/app_repository.py`: `get_by_id`, `list_active`, `list_visible_to`; `app/src/modules/lanzadera/ports/profile_repository.py`: `list_for_app`, `get_by_code`, `create`, `set_active`; `app/src/modules/lanzadera/ports/assignment_repository.py`: `create`, `list_for_user`, `list_for_app`, `effective_permissions`; `app/src/modules/lanzadera/ports/reset_token_repository.py`: `insert`, `find_unused`, `mark_consumed`, `mark_superseded`, `purge_expired`; `app/src/modules/lanzadera/ports/global_admin_repository.py`: `list_all`, `is_global_admin`, `is_there_any`, `grant`, `revoke`; `app/src/modules/lanzadera/ports/credential_hasher.py`: `hash_password`, `verify_password`; `app/src/modules/lanzadera/ports/notification_delivery.py`: `send(to, subject, body) -> None`; `app/src/modules/lanzadera/ports/location.py`: `is_user_in_office`; `app/src/modules/lanzadera/ports/cache.py`: `get`, `set`, `invalidate`; `app/src/modules/lanzadera/ports/audit_log.py`: `append`, `list_for_actor`; `app/src/modules/lanzadera/ports/secret_manager.py`: `get`; `app/src/modules/lanzadera/ports/bootstrap_admin_source.py`: `list_initial_emails`
- [ ] 3.2 **RED**: `tests/lanzadera/auth/test_hash_password.py` — `hash_password` produces Argon2id PHC string with m=65536, t=3, p=4; two calls with same input produce different hashes; **100% coverage** of helper
- [ ] 3.3 **RED**: `tests/lanzadera/auth/test_verify_password.py` — `verify_password` returns True for matching pair, False for wrong; raises `ValueError` on empty/nil hash
- [ ] 3.4 Implement `hash_password` / `verify_password` in `app/src/modules/lanzadera/application/credential_helpers.py` with `argon2-cffi==25.1.0` RFC_9106_LOW_MEMORY profile (DA-2)
- [ ] 3.5 **RED**: `tests/lanzadera/auth/test_issue_reset_token.py` — issues token with `expires_at = now + 24h`; supersedes prior; invokes `NotificationDeliveryPort.send` once; raises `NoGlobalAdminError` when no global admin exists (auth-reset/spec.md §no_global_admin guard)
- [ ] 3.6 **RED**: `tests/lanzadera/auth/test_consume_reset_token.py` — valid token updates password to Argon2id and activates user atomically; rejects unknown/expired/superseded/consumed tokens; rollback if audit fails (DA-11)
- [ ] 3.7 Implement `issue_reset_token` and `consume_reset_token` in `app/src/modules/lanzadera/application/`; token hash via `hashlib.sha256` (internal only, not exported); atomic transaction (DA-4, D90)
- [ ] 3.8 Implement `app/src/modules/lanzadera/application/create_user.py`, `disable_user.py`, `grant_global_admin.py`, `revoke_global_admin.py`, `assign_profile.py`, `list_effective_apps.py`, `audit_append.py`, `bootstrap_global_admins.py`, `set_password.py` (CLI bootstrap, D91)

## Phase 4: Adapters — Driven implementations

- [ ] 4.1 Create `app/src/modules/lanzadera/adapters/postgres/session.py`: `async_session_factory` from `asyncpg` + SQLAlchemy 2.0 Core; bound to `DATABASE_URL` from composition root
- [ ] 4.2 Create `app/src/modules/lanzadera/adapters/postgres/user_repository_pg.py`: implements `UserRepositoryPort`; `dni_encrypted` encrypted via `NationalIdCipher` before INSERT; `update_status` for lockout transitions; `set_password_and_activate` for `consume_reset_token`
- [ ] 4.3 Create `app/src/modules/lanzadera/adapters/postgres/app_repository_pg.py`, `profile_repository_pg.py`, `assignment_repository_pg.py`, `reset_token_repository_pg.py`, `global_admin_repository_pg.py`, `audit_log_pg.py`; `audit_log_pg` appends in the **same transaction** as the calling use case (DA-11)
- [ ] 4.4 Create `app/src/modules/lanzadera/adapters/postgres/tests/` with in-memory fakes for all repo ports; implement contract tests: `test_user_repository_contract.py`, `test_app_repository_contract.py`, `test_profile_repository_contract.py`, `test_assignment_repository_contract.py`
- [ ] 4.5 Create `app/src/modules/lanzadera/adapters/crypto/credential_hasher_argon2id.py`: implements `CredentialHasherPort`; delegates to `application/credential_helpers.py`
- [ ] 4.6 Create `app/src/modules/lanzadera/adapters/crypto/national_id_cipher.py`: AES-GCM via `SecretManagerPort`; `encrypt(plain: str) -> bytes`; `decrypt(ciphertext: bytes) -> str`; key never in logs/CLI args (CA-S4)
- [ ] 4.7 Create `app/src/modules/lanzadera/adapters/notification/mail_queue_table_adapter.py`: implements `NotificationDeliveryPort`; inserts into `mail_outbox` with `status='pending'` (DA-10; G-5 ABIERTO — adapter v1, SMTP real deferred to P20)
- [ ] 4.8 Create `app/src/modules/lanzadera/adapters/location/assume_in_office_adapter.py`: `is_user_in_office -> True` always (DA-9, H12); **RED**: `tests/lanzadera/test_assume_in_office.py` asserts explicit stub return
- [ ] 4.9 Create `app/src/modules/lanzadera/adapters/cache/ttl_cache_adapter.py`: implements `CachePort` using `cachetools.TTLCache`; candidates: `apps.list_active`, `profiles.list_for_app`, `assignments.effective_permissions`; TTLs from env vars; invalidation on mutation
- [ ] 4.10 Create `app/src/modules/lanzadera/adapters/bootstrap/env_admin_source_adapter.py`: reads `GLOBAL_ADMIN_EMAILS`, splits by `;`, returns list; empty if unset (DA-6); **RED**: `tests/lanzadera/test_bootstrap_adapter.py` — idempotency, unset env, audit emission
- [ ] 4.11 Create `app/src/modules/lanzadera/adapters/secrets/env_secret_manager_adapter.py`: `get(key: str) -> str`; reads from `PLATFORM_SECRET_KEY` env var or generates/persists on first boot (MVP); production swapped by Vault adapter without touching domain
- [ ] 4.12 Create `app/src/modules/lanzadera/di/container.py`: composition root wiring all ports to their adapters; FastAPI `Depends()` declarations for every port
- [ ] 4.13 Create `app/src/modules/lanzadera/di/bootstrap.py`: calls `bootstrap_global_admins` on process start; calls `reset_tokens.purge_expired` in background; idempotent (DA-6)

## Phase 5: Delivery — HTTP HTMX + CLI

- [ ] 5.1 Create `app/src/modules/lanzadera/delivery/http/app.py`: FastAPI router `/admin/`; `Annotated[Port, Depends(get_port)]` for every dependency; routes: `POST /login`, `POST /logout`, `GET /admin/users`, `POST /admin/users`, `PATCH /admin/users/{id}/disable`, `POST /admin/assignments`, `GET /admin/apps`, `POST /admin/apps/{id}/activate`, `GET /admin/audit`; `POST /reset` + `GET /reset` for token entry; HTMX fragment responses with `HX-Retarget`; Alpine.js for minimal interactivity
- [ ] 5.2 Create `app/src/modules/lanzadera/delivery/http/templates/async_jinja/` for: `login.html`, `reset_request.html`, `reset_form.html`, `admin/users.html`, `admin/apps.html`, `admin/assignments.html`, `admin/audit.html`; Mistica token CSS vars for brand consistency
- [ ] 5.3 **RED**: `tests/lanzadera/delivery/test_login_view.py` — POST `/login` with wrong password returns 401; POST `/login` with `status=password_reset_required` triggers `issue_reset_token`; fifth bad attempt locks account (D38); **RED**: `tests/lanzadera/delivery/test_reset_view.py` — valid token consumes and activates; invalid/expired/superseded token returns error; empty body rejected
- [ ] 5.4 Create `app/src/modules/lanzadera/delivery/cli/platform_user.py`: subcommands `set-password`, `grant-global-admin`, `revoke-global-admin`, `list-apps`, `assign-profile`; admin-only gate via `is_global_admin` check; destructive commands (`revoke`) prompt explicit confirmation (D26); secrets via env var, never CLI args (CA-S4)
- [ ] 5.5 **RED**: `tests/lanzadera/delivery/cli/test_set_password_cli.py` — `set-password` creates user with Argon2id hash, sets `status='active'`, adds to `global_admins`; idempotent on second run; rejects if caller is not global admin
- [ ] 5.6 **RED**: `tests/lanzadera/delivery/cli/test_grant_revoke_cli.py` — grant then revoke round-trip; `revoke` last admin rejected; non-admin caller rejected

## Phase 6: Tests + Gates — Full suite

- [ ] 6.1 Create `tests/lanzadera/auth/test_no_legacy_compat.py` (DA-13): AST walk over `app/src/modules/lanzadera/` and `app/src/`; fails if `legacy_hash`, `verify_legacy`, `sha256`, `old_password`, `migrate_password` found; runs as normal pytest; **RED first**: add a dummy `legacy_hash = ...` line in a test file and assert it fails
- [ ] 6.2 Create `tests/lanzadera/auth/test_no_legacy_in_auth.py` (DA-13 global scope): same AST walk on full `app/src/`; wrapper for defense-in-depth
- [ ] 6.3 Create `tests/lanzadera/auth/test_lockout_policy.py`: lockout after 5 failures (D38); one-hour duration (D39); global admin unlock resets counter and transitions to `password_reset_required`; global admin notified on every lockout (D40)
- [ ] 6.4 Create `tests/lanzadera/audit/test_same_transaction_audit.py`: mock `AuditLogPort.append` that raises `RuntimeError`; execute `consume_reset_token`; assert `password_hash` and `consumed_at` are unchanged (atomic rollback, DA-11)
- [ ] 6.5 Create `tests/lanzadera/test_secret_manager_contract.py`: `EnvSecretManagerAdapter.get` returns the configured key; missing key raises `KeyError`; key never appears in stdout/logs
- [ ] 6.6 Create `tests/lanzadera/notifications/test_mail_queue_contract.py`: `send` inserts `mail_outbox` row with `status='pending'`; called exactly once per `issue_reset_token`
- [ ] 6.7 Create `tests/lanzadera/global_admins/test_grant_revoke_round_trip.py`, `test_cannot_revoke_last.py`: grant/revoke round-trip; cannot revoke last admin (constraint enforced in `GlobalAdminRepositoryPg.revoke`)
- [ ] 6.8 Create `tests/test_gate_smoke.py`: each `check_*.py` script invoked with a fixture that violates its contract; assert exit `1` and valid JSON envelope (Hard Rule 18 — degenerate-run guard)
- [ ] 6.9 Run full suite: `pytest --cov=app --cov-fail-under=85 -p app.pytest_plugin.coverage_gate`; assert CRITICAL_HELPERS at 100%; assert global ≥ 85%
- [ ] 6.10 Run `make quality-report`; assert all gates pass; publish `quality-report.json`

## Phase 7: Verification + E2E

- [ ] 7.1 Run `docker compose up -d postgres`; `alembic upgrade head`; assert all 6 migrations apply
- [ ] 7.2 E2E: `httpx` test client — POST `/login` with seeded user (status=`password_reset_required`) → receives token via `mail_outbox`; POST `/reset` with token + new password → user becomes `active`; subsequent POST `/login` with new password → 200
- [ ] 7.3 E2E: `gentle-ai platform user set-password admin@empresa.com` → user active → `issue_reset_token` succeeds (no `no_global_admin` error); lockout after 5 failures verified in DB
- [ ] 7.4 E2E: `migrate_from_access.py` run with fixture SHA256 mismatch → aborts with non-zero exit
- [ ] 7.5 Verify no `legacy_hash`, `legacy_password`, `sha256`, `verify_legacy`, `old_password`, `migrate_password` in `app/src/` (DA-13 AST test green)
- [ ] 7.6 Final `make quality-report` clean; `ruff format --check` passes; `mypy app/src/` passes with `enable_error_code = ["ignore-without-code"]`

## Gap Tracking

| Gap | Status | Owner |
|-----|--------|-------|
| G-1 (email normalisation) | **Closed** in DA-3 — `lower(input.email)` | — |
| G-2 (capability set per app) | **Open — product must decide** before 0003 seed; seed `capabilities = {}` for now | Product |
| G-3 (apps fields not migrated) | **Closed** in DA-7 / 0002 — `Pass`, `Comando`, `URLDIrectorioIconoAplicacion` excluded | — |
| G-4 (audit policy for failed admin creation) | **Open** — log at `WARN` severity in MVP; SOC notification deferred | SOC team |
| G-5 (SMTP notification channel) | **Open** — `MailQueueTableAdapter` MVP; SMTP real at P20 | IT |
| G-6 (password expiry policy) | **Open** — not implemented in MVP; `password_reset_required` covers recovery | Product |
| G-7 (SinAcceso weight in menu) | **Open** — stub menu filter returns all visible apps; behaviour decided in UI epic | UI epic |
| G-8 (audit retention periods) | **Open** — MVP: 90 days hot + 1 year total; `AuditRetentionPort` stub | Compliance |

> **Note**: G-2, G-4, G-5, G-6, G-7, G-8 are **non-blocking** for the apply phase. Each is either addressed by a stub/placeholder or deferred to a later release with the decision owner noted.

## Delivery Units — PR Mapping

| PR | Phase(s) | Focus | Size estimate |
|----|----------|-------|---------------|
| PR 1 | Phase 0 | Foundation scaffold + CI gates + pytest plugin + hexagonal layer gate | ~400–500 lines · **size:exception** |
| PR 2 | Phase 1 | Domain entities (pure, framework-free) | ~250–300 lines |
| PR 3a | Phase 2 | Alembic 0001 schema base (users, apps, profiles, assignments, global_admins) | ~280–320 lines |
| PR 3b | Phase 2 | Alembic 0002–0006 (seeds apps/users, reset_tokens, audit, retention) | ~320–360 lines |
| PR 4 | Phase 3 | Application use cases + `CRITICAL_HELPERS` (hash/verify_password, issue/consume_reset_token) | ~250–320 lines |
| PR 5 | Phase 5 | Delivery: HTTP HTMX views + CLI `gentle-ai platform user ...` | ~180–240 lines |
| PR 6a | Phase 4 | Driven adapters — postgres repos + Argon2id crypto + NationalIdCipher | ~280–320 lines |
| PR 6b | Phase 4 | Driven adapters — notification queue + cache + location + BootstrapAdapter | ~300–340 lines |
| PR 7 | Phase 6 | Tests + gates (DA-13 AST pin, lockout, audit transaction, mail queue contract) | ~140–200 lines |
| PR 8 | Phase 7 | E2E verify + integration suite on docker-compose + quality-gate evidence | ~400–500 lines · **size:exception** |

> **Chain strategy** (decided 2026-08-09): `feature-branch-chain` with tracker branch `feat/lanzadera-mvp-tracker`.
>
> **PR stack**:
>
> ```text
> PR 1 (size:exception, base=main)         ──► merge to tracker
>  └─► PR 2 (base=PR 1)                     ──► merge to tracker
>       └─► PR 3a (base=PR 2)               ──► merge to tracker
>             └─► PR 3b (base=PR 3a)        ──► merge to tracker
>                   └─► PR 4 (base=PR 3b)   ──► merge to tracker
>                         └─► PR 5 (base=PR 4)
>                               └─► PR 6a (base=PR 5)
>                                     └─► PR 6b (base=PR 6a)
>                                           └─► PR 7 (base=PR 6b)
>                                                 └─► PR 8 (size:exception, base=PR 7)
> ```
>
> Only `feat/lanzadera-mvp-tracker` merges to `main`; the stack is collapsed at the end.
>
> **Why split PR 3 and PR 6 but exception PR 1 and PR 8**: migrations 0001–0006 and the driven-adapter surface are deep, low-coupling slices that hold up as independent reviewable units (sub-PRs 3a/3b and 6a/6b each stay under 400 lines and keep their own reversible boundary). PR 1 is indivisible: ruff/mypy/pytest plugin, coverage_gate, hexagonal layer gate, package layout and the first `pyproject.toml` all share the same baseline; splitting would leave every sub-PR with a half-configured repo. PR 8 is indivisible for the same reason: the integration suite on docker-compose, the DA-13 final sweep and the evidence bundle for the quality gates only make sense as a single terminal PR.
