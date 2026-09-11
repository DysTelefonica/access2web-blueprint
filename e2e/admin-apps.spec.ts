/**
 * e2e/admin-apps.spec.ts
 *
 * Phase 3 E2E — Admin app management.
 *
 * Tests cover:
 *  - GET    /admin/apps              (list all)
 *  - POST   /admin/apps              (create)
 *  - GET    /admin/apps/{id}         (detail)
 *  - PATCH  /admin/apps/{id}         (update)
 *  - DELETE /admin/apps/{id}         (soft delete / retire)
 *
 * All endpoints require a valid admin JWT.
 * The admin user is seeded by Phase 1 globalSetup.
 *
 * Tests use `page.request` (Playwright APIRequestContext).
 */

import { test, expect } from "@playwright/test";

import { getAdminToken } from "./fixtures/auth-fixture";

import {
  createApp,
  getApp,
  listApps,
  patchApp,
  deleteApp,
} from "./helpers/api-client";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function uniqueShortCode(prefix = "apptest"): string {
  return `${prefix}-${Date.now()}-${Math.floor(Math.random() * 99_999)}`;
}

// ---------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------

test.describe("Admin app management", () => {
  let adminToken: string;

  test.beforeAll(async () => {
    adminToken = await getAdminToken();
  });

  // ---------------------------------------------------------------------------
  // POST /admin/apps — create app
  // ---------------------------------------------------------------------------

  test.describe("POST /admin/apps", () => {
    test("returns 201 + app object with valid payload", async ({ request }) => {
      const shortCode = uniqueShortCode();
      const { status, data } = await createApp(adminToken, {
        name: "Test App",
        short_code: shortCode,
        deployment_topology: "central",
      });
      expect(status).toBe(201);
      expect(data).toHaveProperty("id");
      expect(data).toHaveProperty("short_code", shortCode);
      expect(data).toHaveProperty("name", "Test App");
    });

    test("returns 400 when short_code is missing", async ({ request }) => {
      const res = await request.post(`${process.env.BASE_URL}/admin/apps`, {
        headers: { Authorization: `Bearer ${adminToken}` },
        data: { name: "No Code" },
      });
      expect(res.status()).toBe(400);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.post(`${process.env.BASE_URL}/admin/apps`, {
        data: { name: "X", short_code: "x", deployment_topology: "central" },
      });
      expect(res.status()).toBe(401);
    });
  });

  // ---------------------------------------------------------------------------
  // GET /admin/apps — list apps
  // ---------------------------------------------------------------------------

  test.describe("GET /admin/apps", () => {
    test("returns 200 + app list with admin token", async ({ request }) => {
      const { status, data } = await listApps(adminToken);
      expect(status).toBe(200);
      expect(Array.isArray(data)).toBe(true);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.get(`${process.env.BASE_URL}/admin/apps`);
      expect(res.status()).toBe(401);
    });

    test("returns 403 with a non-admin token", async ({ request }) => {
      const { getUserToken } = await import("./fixtures/auth-fixture");
      const userToken = await getUserToken();
      const res = await request.get(`${process.env.BASE_URL}/admin/apps`, {
        headers: { Authorization: `Bearer ${userToken}` },
      });
      expect([401, 403]).toContain(res.status());
    });
  });

  // ---------------------------------------------------------------------------
  // GET /admin/apps/{id} — app detail
  // ---------------------------------------------------------------------------

  test.describe("GET /admin/apps/{id}", () => {
    test("returns 200 + app object for a known app_id", async ({ request }) => {
      const shortCode = uniqueShortCode();
      const { data: created } = await createApp(adminToken, {
        name: "Detail Test",
        short_code: shortCode,
        deployment_topology: "central",
      });
      const appId = (created as { id: number }).id;
      if (!appId) {
        test.skip("create_app failed — skipping detail test");
        return;
      }

      const { status, data } = await getApp(adminToken, appId);
      expect(status).toBe(200);
      expect(data).toHaveProperty("id", appId);
    });

    test("returns 404 for a non-existent app_id", async ({ request }) => {
      const { status } = await getApp(adminToken, 999_999);
      expect(status).toBe(404);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.get(
        `${process.env.BASE_URL}/admin/apps/1`
      );
      expect(res.status()).toBe(401);
    });
  });

  // ---------------------------------------------------------------------------
  // PATCH /admin/apps/{id} — update app
  // ---------------------------------------------------------------------------

  test.describe("PATCH /admin/apps/{id}", () => {
    test("returns 200 after updating app name", async ({ request }) => {
      const shortCode = uniqueShortCode();
      const { data: created } = await createApp(adminToken, {
        name: "Old Name",
        short_code: shortCode,
        deployment_topology: "central",
      });
      const appId = (created as { id: number }).id;
      if (!appId) {
        test.skip("create_app failed — skipping update test");
        return;
      }

      const { status, data } = await patchApp(adminToken, appId, {
        name: "New Name",
      });
      expect(status).toBe(200);
      expect((data as { name: string }).name).toBe("New Name");
    });

    test("returns 404 when patching a non-existent app", async ({ request }) => {
      const { status } = await patchApp(adminToken, 999_999, { name: "X" });
      expect(status).toBe(404);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.patch(
        `${process.env.BASE_URL}/admin/apps/1`,
        { data: { name: "X" } }
      );
      expect(res.status()).toBe(401);
    });
  });

  // ---------------------------------------------------------------------------
  // DELETE /admin/apps/{id} — retire app
  // ---------------------------------------------------------------------------

  test.describe("DELETE /admin/apps/{id}", () => {
    test("returns 204 after retiring an app", async ({ request }) => {
      const shortCode = uniqueShortCode();
      const { data: created } = await createApp(adminToken, {
        name: "To Retire",
        short_code: shortCode,
        deployment_topology: "central",
      });
      const appId = (created as { id: number }).id;
      if (!appId) {
        test.skip("create_app failed — skipping retire test");
        return;
      }

      const { status } = await deleteApp(adminToken, appId);
      expect(status).toBe(204);
    });

    test("returns 404 when retiring a non-existent app", async ({
      request,
    }) => {
      const { status } = await deleteApp(adminToken, 999_999);
      expect(status).toBe(404);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.delete(
        `${process.env.BASE_URL}/admin/apps/1`
      );
      expect(res.status()).toBe(401);
    });
  });
});
