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
