/**
 * e2e/fixtures/global-setup.ts
 *
 * Playwright 1.64+ requires the file referenced by `globalSetup` to export a
 * SINGLE function via `export default` (no other named exports). The
 * `auth-fixture.ts` module exports several public helpers
 * (`TEST_ADMIN_*`, `getAdminToken`, `getUserToken`) that specs import
 * directly, so it cannot satisfy the single-export rule.
 *
 * This module is a one-line re-export: it pulls `globalSetup` out of
 * `auth-fixture.ts` and exposes it as the only default export. The setup
 * function itself still lives in `auth-fixture.ts`; only the export shape
 * is split so the Playwright 1.64+ validator accepts it.
 */

export { globalSetup as default } from "./auth-fixture";