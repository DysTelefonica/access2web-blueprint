/**
 * e2e/smoke.spec.ts
 *
 * Sanity smoke — the minimum signal that the container is up and the API
 * is reachable before any functional spec runs.
 *
 * This file exists as a standalone entry point that `docker compose up` users
 * can run without needing the full fixture setup (no DB migration, no seed).
 * The CI release gate runs this alongside the full suite.
 *
 * After Phase 5 the full suite is real and the release gate fails on
 * E2E errors. This spec stays as a fast pre-flight for local iteration
 * (e.g. `docker compose up && BASE_URL=http://localhost:8000 npx playwright test e2e/smoke.spec.ts`).
 */

import { test, expect } from "@playwright/test";

test("app starts and /health responds 200", async ({ request }) => {
  const response = await request.get("/health");
  expect(response.status()).toBe(200);
  const body = (await response.json()) as { status: string };
  expect(body.status).toBe("ok");
});

test("API base responds without crashing", async ({ request }) => {
  // A 4xx is fine (the path doesn't exist); 5xx means the server crashed.
  const response = await request.get("/");
  expect(response.status()).toBeLessThan(500);
});
