# End-to-end tests

Playwright suite exercising the published Docker image of `access2web-blueprint`.

## Status

**Phase 1 ✅ merged** (`main`): auth flow — login, logout, /auth/me,
/auth/me/apps, /auth/me/apps/:id/capabilities, POST /auth/revoke.
Wired in `fixtures/auth-fixture.ts` (globalSetup seeds admin + user via
Python+argon2-cffi inside the running container).

**Phase 2 ✅ PR #610 open** (CI green): admin-user CRUD — list users
paginated, create, get, disable, revoke assignment. Routes added in
`admin_routes_users_json.py`; helpers in `helpers/api-client.ts`; spec in
`admin-users.spec.ts` (13 cases).

**Phase 3 ✅ merged**: admin-apps CRUD.

**Phase 4 ✅ merged**: presence/SSE E2E.

**Phase 5 🔄 in-progress** (#608): CLI E2E.

**All phases complete.**

---

The release gate still uses `|| true` — real wiring is issue #602 (umbrella).

## Phases

| Phase | Issue | Spec | Status |
|---|---|---|---|
| F1 | [#604](https://github.com/DysTelefonica/access2web-blueprint/issues/604) | `auth.spec.ts` — login, logout, /auth/me, /auth/me/apps, capabilities, revoke | ✅ done |
| F2 | [#605](https://github.com/DysTelefonica/access2web-blueprint/issues/605) | `admin-users.spec.ts` — list/create/disable/assign | ✅ done |
| F3 | [#606](https://github.com/DysTelefonica/access2web-blueprint/issues/606) | `admin-apps.spec.ts` — CRUD | 🔄 in-progress |
| F4 | [#607](https://github.com/DysTelefonica/access2web-blueprint/issues/607) | `presence.spec.ts` — heartbeat + SSE stream | 🔄 in-progress |
| F5 | [#608](https://github.com/DysTelefonica/access2web-blueprint/issues/608) | `cli.spec.ts` — set-password + platform user CLI | 🔄 in-progress |

Umbrella: [#602](https://github.com/DysTelefonica/access2web-blueprint/issues/602)

## Architecture

```
e2e/
  README.md                    ← this file
  playwright.config.ts         ← globalSetup + baseURL
  smoke.spec.ts                ← placeholder (promoted from original)
  helpers/
    api-client.ts              ← typed fetch wrapper for every API route
  fixtures/
    auth-fixture.ts             ← globalSetup (alembic + seed) + token helpers
  auth.spec.ts                 ← F1: authentication flow (18 test cases)
  admin-users.spec.ts          ← F2: admin user management (pending)
  admin-apps.spec.ts           ← F3: admin app CRUD (pending)
  presence.spec.ts             ← F4: real-time presence + SSE (pending)
  cli.spec.ts                  ← F5: CLI smoke (pending)
```

## Prerequisites resolved by Phase 1

- ✅ DB fixtures — `globalSetup` runs alembic + seeds users via Python inside container
- ✅ Auth secrets — passwords baked into seed script, hashes computed in-container with argon2-cffi
- ✅ Stable selector map — API-only tests use HTTP status + JSON shape, no DOM selectors
- ✅ Schema migration path — `docker exec alembic upgrade head` in `globalSetup`

## Local run

```bash
# 1. Start the app
docker compose up -d

# 2. Install Playwright + Chromium
npx playwright install --with-deps chromium

# 3. Run Phase 1 suite
BASE_URL=http://localhost:8000 npx playwright test e2e/auth.spec.ts
```

## CI / Release gate

The release workflow (`.github/workflows/release.yml::e2e`) runs the full suite
against the just-published image. After Phase 1, the `|| true` wrapper is still
in place — it will be dropped once all 5 phases are complete and the suite is
verified to pass against a real staging run.

Until then: look at the Playwright HTML report (`playwright-report/`) and the
CI logs for signal.
