# Tasks: Lanzadera MVP — primer slice de la plataforma hexagonal

> Change: `lanzadera-mvp` · Project: `access2web-blueprint` · strict_tdd: true · pytest: `pytest --cov=app --cov-fail-under=85`

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Chain strategy | stacked-to-main |
| 400-line budget risk | High for full PR (size:exception claimed) |
| Decision needed before apply | Yes |
| Chained PRs recommended | Yes |

The orchestrator split the change into single-WU stacked PRs (one per
issue). PR #42 is the `feat/042-reset-token-helpers` slice and ships WU A2
in isolation. The full PR exceeds 400 lines because the vertical slice
requires ports + fakes + tests + 2 services + 1 errors module; the
`size:exception` reason is documented in the PR body.

## Work Unit A2 — `feat(auth): add issue_reset_token and consume_reset_token (CRITICAL_HELPERS)`

**Vertical slice**: Pure-domain services for the D90 reset flow.
- `PasswordHasher` Protocol (DA-2, D88) plus the four driven ports the
  services need (`UserRepository`, `ResetTokenRepository`,
  `GlobalAdminRepository`, `NotificationDelivery`, `AuditLog`).
- `issue_reset_token(email, *, now, users, reset_tokens, global_admins,
  notifications, ttl=24h) -> ResetToken`. Honours the no_global_admin
  guard (D90), supersedes prior unconsumed tokens, persists the row,
  emits the raw token via `NotificationDelivery` exactly once.
- `consume_reset_token(token_str, new_password, *, now, hasher, users,
  reset_tokens, audit) -> None`. Atomic single-use. Raises a distinct
  exception for each failure mode (`InvalidResetTokenError`,
  `ExpiredResetTokenError`, `ResetTokenAlreadyUsedError`). Hashes the
  new password via the `PasswordHasher` port.

**Acceptance criteria**:
- [x] RED tests written first; GREEN impl; TDD evidence recorded in
      `openspec/changes/lanzadera-mvp/apply-progress.md`.
- [x] `pytest tests/lanzadera/auth/test_issue_reset_token.py
      tests/lanzadera/auth/test_consume_reset_token.py` passes (24 tests).
- [x] Modules I touch sit at 97-100% coverage (above the 85% floor).
- [x] `ruff check` clean; `mypy --strict` clean; `check_layers.py` clean.
- [x] Tokens are 64 hex chars (BLAKE2b-256), supersession per-user,
      single-use, raised for replay/expired/unknown/superseded.
- [x] No new code introduces `legacy_hash`, `verify_legacy`, `sha256`,
      `old_password`, or `migrate_password` (DA-13).

**Est. lines**: 280 (orchestrator estimate) — actual PR diff is 765
(`size:exception`).

**Test command**: `pytest tests/lanzadera/auth/test_issue_reset_token.py
tests/lanzadera/auth/test_consume_reset_token.py --cov=app`.

**Runtime harness**: full suite `pytest --cov=app --cov-fail-under=69`
runs at 78% coverage; no regression against the 69% baseline.

**Rollback boundary**: revert the seven new files (one ports module,
two services, one errors module, one fakes, two test files).

**Base**: `origin/main` (stacked-to-main).

**Dependencies**: WU D1 (User, ResetToken entities) — already on
`origin/main`.

**Out of scope**: argon2-cffi adapter (PR 46), Postgres adapters (PR 45),
mail-outbox adapter (PR 47), HTMX routes (PR 53), CLI bootstrap (PR 56).

**References**: D90, DA-2, DA-4, DA-11, QC-5.


--------------------------------------------------------------------------------

## Work Unit A2a — DA-7b fixtures (issues #617)

**Vertical slice**: Fixtures sintéticas derivadas de documentación walkthrough
(`walkthrough-G*.json`, `docs/03-aplicaciones/lanzadera/data-model.md`) mediante
`scripts/generate_lanzadera_fixtures.py` (seed=42, determinista).

**Deliverables**:
- `data/fixtures/lanzadera/apps.json` (20 filas)
- `data/fixtures/lanzadera/users.json` (156 filas, sin PII real)
- `data/fixtures/lanzadera/profiles.json` (160 filas)
- `data/fixtures/lanzadera/assignments.json` (622 filas activas)
- `data/fixtures/lanzadera/audit_events.json` (200 filas)
- `data/fixtures/lanzadera/METADATA.json`
- `scripts/generate_lanzadera_fixtures.py`

**Acceptance criteria**:
- [x] Sin Dysflow, sin pyodbc, sin `.accdb` con contraseña
- [x] Generador determinista (seed=42)
- [x] Ruff clean
- [x] Sin PII real
- [x] Cardinalidades verificadas: 20 apps, 156 users, 160 profiles, 622 assignments, 200 events

**Actual PR**: #626 (`feat/617-lanzadera-fixtures`) · `size:exception` (generated fixture data)

**References**: DA-7b, G-3 (cerrado)


## Work Unit A2b — 0002–0007 seed migrations (issues #618–#621)

**Vertical slice**: Cinco migraciones Alembic que siembran los 5 módulos
desde las fixtures de A2a. Test de integración en
`tests/lanzadera/migrations/test_seed_migrations.py`.

**Deliverables**:
- `0003_seed_apps.py` — 20 apps, idempotente
- `0004_seed_profiles.py` — 160 profiles, `capabilities={}` provisional (G-2 ABIERTO)
- `0005_seed_users.py` — 156 users, `dni_encrypted` AES-256-GCM, `password_hash=NULL`
- `0006_seed_assignments.py` — 622 assignments activas, `SinAcceso` rule aplicada en fixture
- `0007_seed_audit.py` — 200 eventos, sin telemetría (DA-11)
- `app/src/modules/lanzadera/domain/legacy_role_map.py` — extendido con `FLAG_TO_CODE`, `LEGACY_CODES`, `resolve_profile_codes`
- `tests/lanzadera/migrations/test_seed_migrations.py` — integración gated `APAP_INTEGRATION_ENABLED=1`

**Acceptance criteria**:
- [x] Ruff clean en las 6 migraciones + domain module
- [x] Cadena de migraciones correcta: 0001 → 0002_expedientes_schema → 0003-0007
- [x] Idempotencia verificada (skip si count ≥ len(fixture))
- [x] G-2 ABIERTO: `capabilities={}` provisional
- [x] Test de integración: `pytest tests/lanzadera/migrations/test_seed_migrations.py`
- [x] `size:exception` para test file (392 líneas — fixture-test data)

**Actual PRs**:
- #627 (`feat/618-seed-apps`) · superseded by #630
- #628 (`feat/619-seed-profiles`) · superseded by #630
- #629 (`feat/620-seed-users`) · superseded by #630
- #630 (`feat/621-seed-audit`) · **base PR con todas las migraciones**

**References**: DA-3, DA-7, DA-7b, DA-11, DA-12, D89, D52, D53, D58, D85, G-2


## Work Unit A2c — Contract tests para los 5 módulos (issue #631)

**Vertical slice**: Tests de contrato para apps, profiles, users, assignments,
audit. Leen fixture JSON directamente — sin Postgres.

**Deliverables**:
- `tests/lanzadera/apps/test_contract.py` (6 tests)
- `tests/lanzadera/profiles/test_contract.py` (7 tests)
- `tests/lanzadera/users/test_contract.py` (7 tests)
- `tests/lanzadera/assignments/test_contract.py` (6 tests)
- `tests/lanzadera/audit/test_contract.py` (6 tests)
- `estado-planificacion-MVP-lanzadera-seeds.html`

**Acceptance criteria**:
- [x] Ruff clean en los 5 archivos
- [x] `pytest -v`: 32/32 passed
- [x] Cada fixture verificada: counts, uniqueness, valores esperados
- [x] HTML con tabla RAG, evidencia TDD RED/GREEN, gaps abiertos

**Actual PR**: #632 (`feat/631-contract-tests`) · 335 líneas

**References**: DA-3, DA-7, DA-11, DA-12, G-2
