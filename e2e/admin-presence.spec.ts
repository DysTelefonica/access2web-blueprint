/**
 * e2e/admin-presence.spec.ts
 *
 * Phase 4 E2E — Presence / SSE.
 *
 * Tests cover:
 *  - POST /presence/heartbeat  (record heartbeat for a user)
 *  - GET  /presence           (list connected users — admin only)
 *  - GET  /presence/stream    (SSE stream — admin only, short timeout)
 *
 * The heartbeat is intentionally ungated (W60 legacy — no JWT required).
 * W62 auth middleware will gate it in the future.
 */

import { test, expect } from "@playwright/test";
import { exec } from "node:child_process";
import { promisify } from "node:util";

import { getAdminToken, getUserToken } from "./fixtures/auth-fixture";

const execAsync = promisify(exec);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** POST /presence/heartbeat — no auth required (W60 legacy). */
async function postHeartbeat(
  baseUrl: string,
  options: { userId?: string; userToken?: string } = {}
): Promise<{ status: number; data?: Record<string, unknown> }> {
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };
  if (options.userId) headers["X-User-ID"] = options.userId;
  if (options.userToken) headers["Authorization"] = `Bearer ${options.userToken}`;

  const res = await fetch(`${baseUrl}/presence/heartbeat`, {
    method: "POST",
    headers,
  });
  let data: Record<string, unknown> | undefined;
  try {
    data = (await res.json()) as Record<string, unknown>;
  } catch {
    // heartbeat returns {} on success
  }
  return { status: res.status, data };
}

/** GET /presence — admin only. */
async function listConnectedUsers(
  adminToken: string,
  baseUrl: string,
  limit = 100
): Promise<{ status: number; data?: unknown[] }> {
  const res = await fetch(`${baseUrl}/presence?limit=${limit}`, {
    headers: { Authorization: `Bearer ${adminToken}` },
  });
  let data: unknown[] | undefined;
  try {
    data = (await res.json()) as unknown[];
  } catch {
    // empty body
  }
  return { status: res.status, data };
}

/**
 * GET /presence/stream — SSE, admin only.
 * Uses curl with a 3-second timeout to get at most one SSE event frame.
 */
async function fetchPresenceStream(
  adminToken: string,
  baseUrl: string,
  timeoutSeconds = 3
): Promise<{ status: number; body: string }> {
  try {
    const { stdout, stderr, code } = await execAsync(
      [
        "curl",
        "--silent",
        "--max-time", String(timeoutSeconds),
        "--header", `Authorization: Bearer ${adminToken}`,
        `${baseUrl}/presence/stream`,
      ].join(" "),
      { timeout: (timeoutSeconds + 2) * 1000 }
    );
    return {
      status: code === 0 ? 200 : 0,
      body: stdout,
    };
  } catch (err) {
    // curl exits non-zero on timeout
    return { status: 0, body: (err as { stdout?: string }).stdout ?? "" };
  }
}

// ---------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------

test.describe("Presence / SSE", () => {
  let adminToken: string;
  let userToken: string;

  test.beforeAll(async () => {
    adminToken = await getAdminToken();
    userToken = await getUserToken();
  });

  // ---------------------------------------------------------------------------
  // POST /presence/heartbeat
  // ---------------------------------------------------------------------------

  test.describe("POST /presence/heartbeat", () => {
    test("returns 200 without auth", async ({ request }) => {
      const { status } = await postHeartbeat(
        process.env.BASE_URL as string
      );
      expect(status).toBe(200);
    });

    test("returns 200 + ok:true with a valid X-User-ID", async ({ request }) => {
      // Grab a known user_id from the user token's JWT sub claim
      // by calling /auth/me first.
      const { data: meData } = await (
        await fetch(`${process.env.BASE_URL}/auth/me`, {
          headers: { Authorization: `Bearer ${userToken}` },
        })
      ).json() as { data?: { id?: string } };

      const userId = meData?.id;
      if (!userId) {
        test.skip("Could not resolve user_id from /auth/me");
        return;
      }

      const { status, data } = await postHeartbeat(
        process.env.BASE_URL as string,
        { userId }
      );
      expect(status).toBe(200);
      expect(data as Record<string, unknown>).toHaveProperty("ok", true);
    });

    test("returns 400 with an invalid X-User-ID", async ({ request }) => {
      const { status } = await postHeartbeat(
        process.env.BASE_URL as string,
        { userId: "not-a-uuid" }
      );
      expect(status).toBe(400);
    });
  });

  // ---------------------------------------------------------------------------
  // GET /presence
  // ---------------------------------------------------------------------------

  test.describe("GET /presence", () => {
    test("returns 200 + list with admin token", async ({ request }) => {
      const { status, data } = await listConnectedUsers(
        adminToken,
        process.env.BASE_URL as string
      );
      expect(status).toBe(200);
      expect(Array.isArray(data)).toBe(true);
    });

    test("returns 401 without a token", async ({ request }) => {
      const res = await request.get(
        `${process.env.BASE_URL}/presence`
      );
      expect(res.status()).toBe(401);
    });

    test("returns 403 with a non-admin token", async ({ request }) => {
      const res = await request.get(
        `${process.env.BASE_URL}/presence`,
        { headers: { Authorization: `Bearer ${userToken}` } }
      );
      expect([401, 403]).toContain(res.status());
    });

    test("returns 200 with limit param", async ({ request }) => {
      const { status, data } = await listConnectedUsers(
        adminToken,
        process.env.BASE_URL as string,
        1
      );
      expect(status).toBe(200);
      expect(Array.isArray(data)).toBe(true);
      expect((data as unknown[]).length).toBeLessThanOrEqual(1);
    });
  });

  // ---------------------------------------------------------------------------
  // GET /presence/stream — SSE
  // ---------------------------------------------------------------------------

  test.describe("GET /presence/stream", () => {
    test("returns 200 with SSE content-type and data lines (admin)", async () => {
      const { status, body } = await fetchPresenceStream(
        adminToken,
        process.env.BASE_URL as string,
        3
      );
      // SSE endpoint responds with 200 and SSE frames
      expect(status).toBe(200);
      // SSE format: "data: <json>\n\n"
      expect(body.trim()).toMatch(/^data: .+$/s);
    });

    test("returns 401 without a token (curl)", async () => {
      try {
        await execAsync(
          [
            "curl",
            "--silent",
            "--max-time", "2",
            "--write-out", "%{http_code}",
            `${process.env.BASE_URL}/presence/stream`,
          ].join(" ")
        );
      } catch (err) {
        // curl exits non-zero on HTTP 401
        const output = (err as { message?: string }).message ?? "";
        expect(output).toMatch(/401/);
      }
    });

    test("returns 401 with a non-admin token (curl)", async () => {
      try {
        await execAsync(
          [
            "curl",
            "--silent",
            "--max-time", "2",
            "--header", `Authorization: Bearer ${userToken}`,
            "--write-out", "%{http_code}",
            `${process.env.BASE_URL}/presence/stream`,
          ].join(" ")
        );
      } catch (err) {
        const output = (err as { message?: string }).message ?? "";
        expect(output).toMatch(/401/);
      }
    });
  });
});
