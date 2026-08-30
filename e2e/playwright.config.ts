/**
 * Playwright configuration for the access2web-blueprint end-to-end suite.
 *
 * Status: PLACEHOLDER. The full E2E tree (DB fixtures, auth secrets, a stable
 * selector map) is tracked as a TODO in `e2e/README.md`. Until that is closed,
 * the only spec is `smoke.spec.ts`, which asserts "the container booted and
 * the FastAPI /health endpoint answers with a non-5xx status".
 *
 * `baseURL` is supplied by the CI job (see
 * `.github/workflows/release.yml::e2e::Run Playwright E2E tests`). Local
 * runs default to http://localhost:8000 so a developer can iterate against
 * `docker compose up app`.
 */
import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: ".",
  // The smoke spec is the only spec that exists today. A wider `*spec.ts`
  // glob keeps room for additional specs without re-editing the config.
  testMatch: /.*spec\.ts$/,
  timeout: 30_000,
  expect: {
    timeout: 5_000,
  },
  fullyParallel: false,
  // CI runs 1 worker against 1 ephemeral container; parallelism only adds
  // contention. A local developer can override with `--workers=N`.
  workers: 1,
  reporter: process.env.CI ? "list" : "list",
  use: {
    baseURL: process.env.BASE_URL ?? "http://localhost:8000",
    trace: "retain-on-failure",
  },
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] },
    },
  ],
});
