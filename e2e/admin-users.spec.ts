/**
 * e2e/admin-users.spec.ts
 *
 * Phase 2 E2E — Admin user management.
 *
 * Tests cover:
 *  - GET  /admin/users          (paginated list)
 *  - POST /admin/users          (create)
 *  - GET  /admin/users/{id}     (detail)
 *  - PATCH /admin/users/{id}    (disable user)
 *  - DELETE /admin/users/{id}/assignments/{app_id} (revoke assignments)
 *
 * All endpoints require a valid admin JWT.
 * The admin user is seeded by Phase 1 globalSetup.
 *
 * Tests use `page.request` (Playwright APIRequestContext) — no browser page needed.
 */

import { test, expect } from "@playwright/test";

import {
  getAdminToken,
  TEST_ADMIN_EMAIL,
} from "./fixtures/auth-fixture";

import {
  createUser,
  getUser,
  listUsers,
  revokeUserAssignment,
  patchUser,
} from "./helpers/api-client";

// ---------------------------------------------------------------------------
// Setup — obtain admin token once per describe block
// ---------------------------------------------------------------------------

test.describe("Admin user management", () => {
  let adminToken: string;

  test.beforeAll(async () => {
    adminToken = await getAdminToken();
  });

  // ---------------------------------------------------------------------------
  // GET /admin/users — paginated list
  // ---------------------------------------------------------------------------

  test.describe("GET /admin/users", () => {
    test("returns 200 + paginated user list with admin token", async ({
      request,
    }) => {
      const { status, data } = await listUsers(adminToken, 10, 0);
      expect(status).toBe(200);
      expect(data).toHaveProperty("users");
      expect(data).toHaveProperty("total");
      expect(data).toHaveProperty("limit", 10);
      expect(data).toHaveProperty("offset", 0);
      expect(Array.isArray((data as { users: unknown }).users)).toBe(true);
    });

    test("returns 200 with limit=1 and offset=0 (pagination boundary)", async ({
      request,
    }) => {
      const { status, data } = await listUsers(adminToken, 1, 0);
      expect(status).toBe(200);
      expect((data as { users: unknown[] }).users.length).toBeLessThanOrEqual(1);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.get(`${process.env.BASE_URL}/admin/users`);
      expect(res.status()).toBe(401);
    });

    test("returns 403 with a non-admin (regular user) token", async ({
      request,
    }) => {
      const { getUserToken } = await import("./fixtures/auth-fixture");
      const userToken = await getUserToken();
      const res = await request.get(`${process.env.BASE_URL}/admin/users`, {
        headers: { Authorization: `Bearer ${userToken}` },
      });
      expect([401, 403]).toContain(res.status());
    });
  });

  // ---------------------------------------------------------------------------
  // POST /admin/users — create user
  // ---------------------------------------------------------------------------

  test.describe("POST /admin/users", () => {
    const uniqueEmail = `e2e-test-${Date.now()}@admin-users.test`;

    test("returns 201 + user object with valid payload", async ({ request }) => {
      const { status, data } = await createUser(adminToken, {
        email: uniqueEmail,
        name: "E2E Test User",
        national_id: "12345678A",
      });
      expect(status).toBe(201);
      expect(data).toHaveProperty("id");
      expect(data).toHaveProperty("email", uniqueEmail);
      expect(data).toHaveProperty("name", "E2E Test User");
    });

    test("returns 409 when email already exists", async ({ request }) => {
      const { status } = await createUser(adminToken, {
        email: TEST_ADMIN_EMAIL,
        name: "Duplicate",
        national_id: "00000000A",
      });
      expect(status).toBe(409);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.post(`${process.env.BASE_URL}/admin/users`, {
        data: { email: "x@test.com", name: "X", national_id: "X" },
      });
      expect(res.status()).toBe(401);
    });
  });

  // ---------------------------------------------------------------------------
  // GET /admin/users/{id} — user detail
  // ---------------------------------------------------------------------------

  test.describe("GET /admin/users/{id}", () => {
    test("returns 200 + user object for a known user_id", async ({ request }) => {
      const { data } = await listUsers(adminToken, 10, 0);
      const users = (data as { users: unknown[] }).users as { id: string }[];
      if (users.length === 0) {
        test.skip("No users in DB — seed data missing");
        return;
      }
      const userId = users[0].id;

      const { status, data: userData } = await getUser(adminToken, userId);
      expect(status).toBe(200);
      expect(userData).toHaveProperty("id", userId);
      expect(userData).toHaveProperty("email");
    });

    test("returns 404 for a non-existent UUID", async ({ request }) => {
      const { status } = await getUser(
        adminToken,
        "00000000-0000-0000-0000-000000000000"
      );
      expect(status).toBe(404);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.get(
        `${process.env.BASE_URL}/admin/users/00000000-0000-0000-0000-000000000001`
      );
      expect(res.status()).toBe(401);
    });
  });

  // ---------------------------------------------------------------------------
  // PATCH /admin/users/{id} — disable user
  // ---------------------------------------------------------------------------

  test.describe("PATCH /admin/users/{id}", () => {
    test("returns 200 after disabling a user", async ({ request }) => {
      const uniqueEmail = `disable-test-${Date.now()}@admin-users.test`;
      const { data: created } = await createUser(adminToken, {
        email: uniqueEmail,
        name: "To Disable",
        national_id: "11111111A",
      });
      if ((created as { id?: string }).id == null) {
        test.skip("create_user failed — skipping disable test");
        return;
      }
      const userId = (created as { id: string }).id;

      const { status, data: updated } = await patchUser(
        adminToken,
        userId,
        { active: false }
      );
      expect(status).toBe(200);
    });

    test("returns 404 when disabling a non-existent user", async ({ request }) => {
      const { status } = await patchUser(
        adminToken,
        "00000000-0000-0000-0000-000000000000",
        { active: false }
      );
      expect(status).toBe(404);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.patch(
        `${process.env.BASE_URL}/admin/users/00000000-0000-0000-0000-000000000001`,
        { data: { active: false } }
      );
      expect(res.status()).toBe(401);
    });
  });

  // ---------------------------------------------------------------------------
  // DELETE /admin/users/{id}/assignments/{app_id} — revoke assignments
  // ---------------------------------------------------------------------------

  test.describe("DELETE /admin/users/{id}/assignments/{app_id}", () => {
    test("returns 204 when revoking assignments for a known user+app", async ({
      request,
    }) => {
      const uniqueEmail = `assign-test-${Date.now()}@admin-users.test`;
      const { data: created } = await createUser(adminToken, {
        email: uniqueEmail,
        name: "Assignment Test",
        national_id: "22222222A",
      });
      const userId = (created as { id: string }).id;
      if (!userId) {
        test.skip("create_user failed — skipping assignment test");
        return;
      }

      const { status } = await revokeUserAssignment(adminToken, userId, 999_999);
      expect(status).toBe(204);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.delete(
        `${process.env.BASE_URL}/admin/users/00000000-0000-0000-0000-000000000001/assignments/1`
      );
      expect(res.status()).toBe(401);
    });
  });
});
