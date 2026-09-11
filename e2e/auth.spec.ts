/**
 * e2e/auth.spec.ts
 *
 * Phase 1 E2E — Authentication flow.
 *
 * Covers:
 *  - POST /auth/login (success + wrong password)
 *  - POST /auth/logout (success + unauthenticated)
 *  - GET  /auth/me (authenticated + unauthenticated)
 *  - GET  /auth/me/apps (authenticated + unauthenticated)
 *  - GET  /auth/me/apps/:id/capabilities (authenticated + unauthenticated)
 *  - POST /auth/me/apps/:id/capabilities/revoke (authenticated + unauthenticated)
 *
 * Prerequisites:
 *  - globalSetup (fixtures/auth-fixture.ts) ran alembic upgrade head and seeded
 *    TEST_ADMIN_EMAIL + TEST_USER_EMAIL with known Argon2id password hashes.
 *  - Both users can log in and receive a JWT.
 *
 * All tests use Playwright's `page.request` API — no browser page needed for
 * HTTP API assertions.
 */

import { test, expect } from "@playwright/test";

import {
  getAdminToken,
  getUserToken,
  TEST_ADMIN_EMAIL,
  TEST_USER_EMAIL,
} from "./fixtures/auth-fixture";

import {
  getMyApps,
  getMyCapabilities,
  login,
  logout,
  revokeAssignment,
} from "./helpers/api-client";

// ---------------------------------------------------------------------------
// Login
// ---------------------------------------------------------------------------

test.describe("POST /auth/login", () => {
  test("returns 200 + token + session_id with valid credentials", async ({
    request,
  }) => {
    const { status, data } = await login(TEST_USER_EMAIL, "AliceE2E!Sec0r3Pwd");
    expect(status).toBe(200);
    expect(data).toHaveProperty("token");
    expect(data).toHaveProperty("session_id");
    expect(typeof (data as { token: unknown }).token).toBe("string");
    expect((data as { session_id: unknown }).session_id).toMatch(
      /^[0-9a-f-]{36}$/ // UUID format
    );
  });

  test("returns 200 with admin credentials", async ({ request }) => {
    const { status, data } = await login(TEST_ADMIN_EMAIL, "AdminE2E!Sec0r3Pwd");
    expect(status).toBe(200);
    expect(data).toHaveProperty("token");
  });

  test("returns 401 with wrong password", async ({ request }) => {
    const { status, data } = await login(TEST_USER_EMAIL, "wrongpassword123");
    expect(status).toBe(401);
    expect(data).toHaveProperty("detail");
  });

  test("returns 401 with non-existent email", async ({ request }) => {
    const { status } = await login(
      "nobody@e2e.lanzadera.test",
      "anypassword"
    );
    expect(status).toBe(401);
  });

  test("returns 422 with missing email field", async ({ request }) => {
    const res = await request.post(`${process.env.BASE_URL}/auth/login`, {
      data: { password: "test" },
    });
    expect(res.status()).toBe(422);
  });
});

// ---------------------------------------------------------------------------
// Logout
// ---------------------------------------------------------------------------

test.describe("POST /auth/logout", () => {
  test("returns 204 with a valid token", async ({ request }) => {
    const token = await getUserToken();
    const { status } = await logout(token);
    expect(status).toBe(204);
  });

  test("returns 401 when called without a token", async ({ request }) => {
    const res = await request.post(`${process.env.BASE_URL}/auth/logout`);
    // Without Authorization header, middleware leaves state empty → 401
    expect(res.status()).toBe(401);
  });

  test("returns 401 when token is malformed", async ({ request }) => {
    const res = await request.post(`${process.env.BASE_URL}/auth/logout`, {
      headers: { Authorization: "Bearer not-a-real-token" },
    });
    // Invalid JWT → JWTError → middleware leaves state empty → 401
    expect(res.status()).toBe(401);
  });
});

// ---------------------------------------------------------------------------
// GET /auth/me
// ---------------------------------------------------------------------------

test.describe("GET /auth/me", () => {
  test("returns 200 + user identity when authenticated", async ({
    request,
  }) => {
    const token = await getUserToken();
    const res = await request.get(`${process.env.BASE_URL}/auth/me`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    expect(res.status()).toBe(200);
    const body = (await res.json()) as {
      user_id: string;
      email: string;
      name: string;
      apps: unknown[];
    };
    expect(body).toHaveProperty("user_id");
    expect(body).toHaveProperty("email");
    expect(body.email).toBe(TEST_USER_EMAIL);
  });

  test("returns 401 when no token is provided", async ({ request }) => {
    const res = await request.get(`${process.env.BASE_URL}/auth/me`);
    expect(res.status()).toBe(401);
  });

  test("returns 401 when token is expired or invalid", async ({
    request,
  }) => {
    const res = await request.get(`${process.env.BASE_URL}/auth/me`, {
      headers: { Authorization: "Bearer eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJ0ZXN0In0.fake" },
    });
    expect(res.status()).toBe(401);
  });

  test("returns admin identity with admin token", async ({ request }) => {
    const token = await getAdminToken();
    const res = await request.get(`${process.env.BASE_URL}/auth/me`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    expect(res.status()).toBe(200);
    const body = (await res.json()) as { email: string };
    expect(body.email).toBe(TEST_ADMIN_EMAIL);
  });
});

// ---------------------------------------------------------------------------
// GET /auth/me/apps
// ---------------------------------------------------------------------------

test.describe("GET /auth/me/apps", () => {
  test("returns 200 + app list when authenticated", async ({ request }) => {
    const token = await getUserToken();
    const { status, data } = await getMyApps(token);
    expect(status).toBe(200);
    expect(data).toHaveProperty("user_id");
    expect(data).toHaveProperty("apps");
    expect(Array.isArray((data as { apps: unknown }).apps)).toBe(true);
  });

  test("returns 401 when no token is provided", async ({ request }) => {
    const res = await request.get(`${process.env.BASE_URL}/auth/me/apps`);
    expect(res.status()).toBe(401);
  });
});

// ---------------------------------------------------------------------------
// GET /auth/me/apps/:app_id/capabilities
// ---------------------------------------------------------------------------

test.describe("GET /auth/me/apps/:app_id/capabilities", () => {
  test("returns 200 + capabilities array for a known app_id", async ({
    request,
  }) => {
    const token = await getUserToken();
    // First get the list of apps
    const { data: appsData } = await getMyApps(token);
    const apps = (appsData as { apps: { app_id: number }[] }).apps;
    if (apps.length === 0) {
      test.skip("No apps assigned to test user — skip (expected in early E2E)");
      return;
    }
    const appId = apps[0].app_id;

    const { status, data } = await getMyCapabilities(token, appId);
    expect(status).toBe(200);
    expect(data).toHaveProperty("user_id");
    expect(data).toHaveProperty("app_id");
    expect(data).toHaveProperty("capabilities");
    expect(Array.isArray((data as { capabilities: unknown }).capabilities)).toBe(
      true
    );
  });

  test("returns 200 with empty capabilities when user has no assignment for that app", async ({
    request,
  }) => {
    const token = await getUserToken();
    // Use a very large app_id that won't exist
    const { status, data } = await getMyCapabilities(token, 999_999_999);
    expect(status).toBe(200);
    expect((data as { capabilities: string[] }).capabilities).toEqual([]);
  });

  test("returns 401 when no token is provided", async ({ request }) => {
    const res = await request.get(
      `${process.env.BASE_URL}/auth/me/apps/1/capabilities`
    );
    expect(res.status()).toBe(401);
  });
});

// ---------------------------------------------------------------------------
// POST /auth/me/apps/:app_id/capabilities/revoke
// ---------------------------------------------------------------------------

test.describe("POST /auth/me/apps/:app_id/capabilities/revoke", () => {
  test("returns 204 when authenticated", async ({ request }) => {
    const token = await getUserToken();
    // Revoke for a non-existent app — idempotent, returns 204 regardless
    const { status } = await revokeAssignment(token, 999_999_999);
    expect(status).toBe(204);
  });

  test("returns 401 when no token is provided", async ({ request }) => {
    const res = await request.post(
      `${process.env.BASE_URL}/auth/me/apps/1/capabilities/revoke`
    );
    expect(res.status()).toBe(401);
  });

  test("revoke + re-fetch confirms the assignment is gone", async ({
    request,
  }) => {
    // This test requires a user with an app assignment.
    // It revokes the assignment and verifies subsequent GET returns empty.
    // Skip if the test user has no assignments (early E2E before app CRUD).
    const token = await getUserToken();
    const { data: appsData } = await getMyApps(token);
    const apps = (appsData as { apps: { app_id: number }[] }).apps;
    if (apps.length === 0) {
      test.skip("No apps assigned to test user — skip (assign via admin API first)");
      return;
    }

    const appId = apps[0].app_id;

    // Revoke
    const revokeRes = await revokeAssignment(token, appId);
    expect(revokeRes.status).toBe(204);

    // Re-fetch
    const { status, data } = await getMyCapabilities(token, appId);
    expect(status).toBe(200);
    expect((data as { capabilities: string[] }).capabilities).toEqual([]);
  });
});
