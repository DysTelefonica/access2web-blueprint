# End-to-end tests

Playwright suite that exercises the published Docker image of `access2web-blueprint`.
The release workflow runs these against `ghcr.io/access2web-blueprint/access2web-blueprint:vX.Y.Z`
right after the image is pushed, before publishing the GitHub release.

## Status

**Placeholder.** This directory exists so the release gate has a non-blocking
plumbing probe while the real E2E suite is built. The current `smoke.spec.ts`
asserts that the container booted and `/health` answers with a non-5xx status.
The release job wraps the Playwright invocation with `|| true` by design — see
`Smoke test — the only spec the placeholder E2E tree ships with.` in
`smoke.spec.ts` for the rationale and the `Run Playwright E2E tests` step in
`.github/workflows/release.yml::e2e` for the exact plumbing.

## Why this is a placeholder

A real E2E gate needs four things the project does not have today:

1. **DB fixtures.** Tests against the admin surface need a seeded sqlite +
   migrations applied. The image ships with migrations baked in, but no
   fixture loader is wired to the container.
2. **Auth secrets.** The auth flow needs a known-seed user, a known-seed
   password hash, and a JWT signing secret. None of those is currently
   configurable for the test container without leaking production paths.
3. **A stable selector map.** The admin UI is in active design; pinning
   selectors today would mean re-authorising every spec on every UI change.
4. **A schema migration path.** The release gate pulls the published image
   verbatim; running a test migration as part of E2E means either baking
   migrations into the test entrypoint or shipping a parallel migrator.

Until those four are built, the only honest assertion is "the container
booted and /health answered".

## Files

```
e2e/
  README.md             # this file
  playwright.config.ts  # baseURL via $BASE_URL, chromium project, 1 worker
  smoke.spec.ts         # placeholder: GET /health non-5xx; TODO marker test
```

## Local run

```bash
# 1. Boot the app (any way you like; the suite only needs /health on :8000)
docker compose up app

# 2. Install playwright + chromium
npx playwright install --with-deps chromium

# 3. Run the suite against http://localhost:8000
npx playwright test e2e/ --config=e2e/playwright.config.ts
```

The release job sets `--base-url=http://localhost:8000` and points
`baseURL` at the same port for the ephemeral container — see
`.github/workflows/release.yml::e2e`.

## TODO: real E2E

When the four prerequisites above land, replace `smoke.spec.ts` with:

- `health.spec.ts` — the current smoke, promoted out of placeholder
- `auth.spec.ts` — login, refresh, role-bound routes
- `admin-apps.spec.ts` — `/admin/apps` create + list / disable
- `walkthrough.spec.ts` — public flow surfaced by the v1.4 walkthrough

At the same time, drop the `|| true` from the release job's Playwright step
so the E2E suite becomes a real release gate. The job's purpose is to catch
the cases that the unit-test boundary cannot — wrong image tag, wrong port,
wrong env defaults — so wiring it up tight matters as soon as the assertions
are real.
