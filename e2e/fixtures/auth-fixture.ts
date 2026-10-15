/**
 * e2e/fixtures/auth-fixture.ts
 *
 * Global Playwright fixtures for the Lanzadera E2E suite.
 *
 * Responsibilities:
 *  1. Seed two known users (admin and regular) into the real `lanzadera`
 *     schema, after the CI runner has applied `alembic upgrade head`.
 *     The migrations are not run here: they live in ci.yml, next to the
 *     service that hosts the database.
 *  2. Provide `getAdminToken()` and `getUserToken()` to every test via module
 *     exports — each spec calls them per test or caches at describe-level.
 *
 * Runtime contract (read from env):
 *  - `TEST_CONTAINER_NAME`  — defaults to `test-app` (matches the release workflow).
 *  - `DATABASE_URL`         — `postgresql+asyncpg://...` (the same value the
 *                             app container receives).
 *
 * Usage in a spec:
 *
 *   import { test } from "@playwright/test";
 *   import { getAdminToken, getUserToken } from "./fixtures/auth-fixture";
 *
 *   test("authenticated me endpoint", async ({ request }) => {
 *     const token = await getUserToken();
 *     const res = await request.get("/auth/me", {
 *       headers: { Authorization: `Bearer ${token}` },
 *     });
 *     expect(res.ok()).toBe(true);
 *   });
 */

import { type FullConfig } from "@playwright/test";

import { login } from "../helpers/api-client";

// ---------------------------------------------------------------------------
// Known credentials — stable across all E2E runs
// ---------------------------------------------------------------------------

export const TEST_ADMIN_EMAIL = "admin@e2e.lanzadera.test";
export const TEST_ADMIN_PASSWORD = "AdminE2E!Sec0r3Pwd";
export const TEST_ADMIN_NAME = "E2E Admin";

export const TEST_USER_EMAIL = "alice@e2e.lanzadera.test";
export const TEST_USER_PASSWORD = "AliceE2E!Sec0r3Pwd";
export const TEST_USER_NAME = "Alice E2E";

// ---------------------------------------------------------------------------
// globalSetup — run once per worker before any test
// ---------------------------------------------------------------------------

export async function globalSetup(_config: FullConfig): Promise<void> {
  const containerName = process.env.TEST_CONTAINER_NAME ?? "test-app";

  // 1. Migrations are applied by the CI step that runs before the app
  //    starts (ci.yml, `alembic upgrade head` on the runner). They cannot
  //    run in the container: 0003 reads `data/fixtures/` from the repo
  //    root, which the image does not contain.

  // 2. Seed the two test users into the real schema. The script runs in the
  //    container, where argon2-cffi and asyncpg are installed, so the hashes
  //    come from the same hasher family the app uses for login.
  //
  //    `dni_encrypted` is NOT NULL in the schema. The login flow never reads
  //    it, so the test rows carry a fixed placeholder instead of an AES-GCM
  //    ciphertext. Status is `active`, because login rejects any other state.
  //    The insert is idempotent (ON CONFLICT on the fixed UUID).
  const seedScript = `
import asyncio, os, uuid
import asyncpg
from argon2 import PasswordHasher

dsn = os.environ["DATABASE_URL"].replace("postgresql+asyncpg://", "postgresql://", 1)
h = PasswordHasher()
users = [
    (uuid.UUID("00000000-0000-0000-0000-000000000001"), "${TEST_ADMIN_EMAIL}", "${TEST_ADMIN_NAME}", h.hash("${TEST_ADMIN_PASSWORD}")),
    (uuid.UUID("00000000-0000-0000-0000-000000000002"), "${TEST_USER_EMAIL}", "${TEST_USER_NAME}", h.hash("${TEST_USER_PASSWORD}")),
]

async def main():
    conn = await asyncpg.connect(dsn)
    try:
        for uid, email, name, pw_hash in users:
            await conn.execute(
                """
                INSERT INTO lanzadera.users (id, email, name, dni_encrypted, password_hash, status)
                VALUES ($1, $2, $3, $4, $5, 'active')
                ON CONFLICT (id) DO NOTHING
                """,
                uid, email, name, b"e2e-placeholder", pw_hash,
            )
        rows = await conn.fetch(
            "SELECT email, status::text AS status FROM lanzadera.users WHERE email IN ($1, $2)",
            "${TEST_ADMIN_EMAIL}", "${TEST_USER_EMAIL}",
        )
        print(f"Seeded users: {[(r['email'], r['status']) for r in rows]}")
    finally:
        await conn.close()

asyncio.run(main())
`;

  const seedResult = await runInContainer(containerName, [
    "python", "-c", seedScript,
  ], { cwd: "/app" });

  if (seedResult.exitCode !== 0) {
    throw new Error(
      `e2e:globalSetup — seed script failed:\nstdout: ${seedResult.stdout}\nstderr: ${seedResult.stderr}`
    );
  }
  console.log(`[e2e:globalSetup] ${seedResult.stdout.trim()}`);
}

// ---------------------------------------------------------------------------
// Token helpers — call login() against the live container
// ---------------------------------------------------------------------------

/**
 * Returns a valid admin JWT. Each call hits the live /auth/login endpoint.
 * Specs that run many tests in one describe block should call this once and
 * share the token to avoid N login round-trips per spec run.
 */
export async function getAdminToken(): Promise<string> {
  const { status, token } = await login(TEST_ADMIN_EMAIL, TEST_ADMIN_PASSWORD);
  if (status !== 200 || !token) {
    throw new Error(
      `e2e:getAdminToken — login returned ${status}, token=${token}. ` +
      `Is the admin user seeded? Check globalSetup logs above.`
    );
  }
  return token;
}

/**
 * Returns a valid regular-user JWT.
 */
export async function getUserToken(): Promise<string> {
  const { status, token } = await login(TEST_USER_EMAIL, TEST_USER_PASSWORD);
  if (status !== 200 || !token) {
    throw new Error(
      `e2e:getUserToken — login returned ${status}, token=${token}. ` +
      `Is the regular user seeded? Check globalSetup logs above.`
    );
  }
  return token;
}

// ---------------------------------------------------------------------------
// Docker exec helper
// ---------------------------------------------------------------------------

interface ContainerResult {
  exitCode: number;
  stdout: string;
  stderr: string;
}

async function runInContainer(
  containerName: string,
  cmd: string[],
  options: { cwd?: string; env?: Record<string, string> } = {}
): Promise<ContainerResult> {
  const { spawn } = await import("node:child_process");
  const cwd = options.cwd ?? "/app";

  return new Promise((resolve) => {
    const args = ["exec", "-w", cwd];
    for (const [k, v] of Object.entries(options.env ?? {})) {
      args.push("-e", `${k}=${v}`);
    }
    args.push(containerName, ...cmd);

    const proc = spawn("docker", args);
    let stdout = "";
    let stderr = "";
    proc.stdout?.on("data", (d) => (stdout += d.toString()));
    proc.stderr?.on("data", (d) => (stderr += d.toString()));
    proc.on("close", (code) =>
      resolve({ exitCode: code ?? 0, stdout, stderr })
    );
    proc.on("error", (err) =>
      resolve({ exitCode: 127, stdout, stderr: err.message })
    );
  });
}
