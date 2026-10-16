/**
 * e2e/playwright.config.ts
 *
 * Playwright configuration for the access2web-blueprint end-to-end suite.
 *
 * ## Runtime model
 *
 * Tests run against a **running Docker container** started by the release
 * workflow (`.github/workflows/release.yml::e2e`). The container exposes
 * the app on port 8000; `BASE_URL` and `TEST_CONTAINER_NAME` are injected
 * by the workflow as env vars.
 *
 * The release gate runs all specs (auth, admin-users, admin-apps,
 * admin-presence, cli, smoke) against the just-published image.
 * Failures block the release (no `|| true`).
 *
 * ## globalSetup
 *
 * `fixtures/auth-fixture.ts::globalSetup` runs once per worker before any
 * test. It:
 *   1. Runs `alembic upgrade head` inside the container.
 *   2. Seeds `admin@e2e.lanzadera.test` and `alice@e2e.lanzadera.test`
 *      with pre-computed Argon2id password hashes (Python inside the container).
 *
 * ## Secrets
 *
 * `TEST_CONTAINER_NAME` and `DATABASE_URL` may be set as env vars; both
 * default to the values used by the release workflow (`test-app`,
 * `sqlite+aiosqlite:///./test.db`).
 *
 * ## Local run
 *
 * ```bash
 * # 1. Start the app
 * docker compose up -d
 *
 * # 2. Run migrations and seed (optional — globalSetup does this automatically)
 * docker compose exec app python -m alembic upgrade head
 *
 * # 3. Install Playwright + Chromium
 * npx playwright install --with-deps chromium
 *
 * # 4. Run the suite
 * BASE_URL=http://localhost:8000 npx playwright test e2e/
 * ```
 */

import { defineConfig, devices } from "@playwright/test";

export const BASE_URL = process.env.BASE_URL ?? "http://localhost:8000";

export default defineConfig({
  testDir: ".",
  testMatch: /.*spec\.ts$/,
  timeout: 30_000,
  expect: { timeout: 5_000 },
  fullyParallel: false,
  // One worker: avoids contention when running against a single ephemeral container.
  // Override locally with `--workers=N`.
  workers: 1,
  reporter: process.env.CI ? "list" : "list",
  use: {
    baseURL: BASE_URL,
    trace: "retain-on-failure",
  },
  // Playwright 1.64+ requires `globalSetup` to be a string path AND the
  // pointed-to module must export a single function via `export default`
  // (no other named exports). `auth-fixture.ts` carries several public
  // helpers (TEST_ADMIN_*, getAdminToken, getUserToken) that specs
  // import directly, so it cannot satisfy the single-export rule.
  // `fixtures/global-setup.ts` is the single-export sibling: it imports
  // only the credentials it needs, keeps the docker-exec helper private,
  // and re-exports nothing else.
  globalSetup: "./fixtures/global-setup.ts",
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] },
    },
  ],
});
