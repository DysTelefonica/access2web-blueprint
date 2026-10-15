/**
 * e2e/fixtures/auth-fixture.ts
 *
 * Global Playwright fixtures for the Lanzadera E2E suite.
 *
 * Responsibilities:
 *  1. Run Alembic migrations against the live container before any test fires.
 *     This ensures the DB schema is up-to-date regardless of how the image was
 *     built (the Dockerfile does not bake in migrations).
 *  2. Seed known admin and regular users directly into SQLite.
 *     The seed script runs Python inside the container (where argon2-cffi is
 *     available) so it can hash passwords with the real hasher.
 *  3. Provide `getAdminToken()` and `getUserToken()` to every test via module
 *     exports — each spec calls them per test or caches at describe-level.
 *
 * Runtime contract (read from env):
 *  - `TEST_CONTAINER_NAME`  — defaults to `test-app` (matches the release workflow).
 *  - `DATABASE_URL`         — defaults to `sqlite+aiosqlite:///./test.db`.
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

  // 2. Seed test users — Python script runs inside the container where
  //    argon2-cffi is available, so we get real hashes for the test passwords.
  //
  //    argon2-cffi 23+ renamed `PasswordHash` to `PasswordHasher`. The
  //    container pins argon2-cffi==25.1.0 (see app/pyproject.toml); the
  //    app's own hasher (`credential_hasher_argon2id.py`) also uses
  //    `PasswordHasher`. The old `PasswordHash` import worked against
  //    an earlier binding; the PR's release-yml / nightly runs hit the
  //    `ImportError` on the new image.
  //
  //    The DB path is taken from $DATABASE_URL (same env var alembic
  //    consumes) and resolved to an absolute path with `os.path.abspath`
  //    when relative. With the migrations step now pointing at
  //    `/app/migrations/alembic.ini` and the container's WORKDIR=/app,
  //    alembic creates the file at `/app/test.db`; the seed opens the
  //    same path so both writes agree.
  const seedScript = `
import os, sqlite3, sys, uuid
from argon2 import PasswordHasher

db_url = os.environ.get("DATABASE_URL", "sqlite+aiosqlite:///./test.db")
db_path = db_url.split("///", 1)[1]
if db_path and not db_path.startswith("/"):
    db_path = os.path.abspath(db_path)
print(f"[e2e:seed] cwd={os.getcwd()} db_url={db_url} db_path={db_path}")
if not os.path.exists(db_path):
    sys.exit(f"seed: database file not found at {db_path}")

# Generate real Argon2id hashes using the container's argon2-cffi
h = PasswordHasher()
admin_hash = h.hash("${TEST_ADMIN_PASSWORD}")
user_hash  = h.hash("${TEST_USER_PASSWORD}")

conn = sqlite3.connect(db_path)
cur = conn.cursor()

admin_id = str(uuid.UUID("00000000-0000-0000-0000-000000000001"))
user_id  = str(uuid.UUID("00000000-0000-0000-0000-000000000002"))

# Insert admin — use INSERT OR IGNORE so re-runs are idempotent
cur.execute("""
    INSERT OR IGNORE INTO users (id, email, name, password_hash, active,
                                 created_at, updated_at, legacy_hash)
    VALUES (?, ?, ?, ?, 1, datetime('now'), datetime('now'), NULL)
""", (admin_id, "${TEST_ADMIN_EMAIL}", "${TEST_ADMIN_NAME}", admin_hash))

# Insert regular user
cur.execute("""
    INSERT OR IGNORE INTO users (id, email, name, password_hash, active,
                                 created_at, updated_at, legacy_hash)
    VALUES (?, ?, ?, ?, 1, datetime('now'), datetime('now'), NULL)
""", (user_id, "${TEST_USER_EMAIL}", "${TEST_USER_NAME}", user_hash))

conn.commit()

# Verify both inserted
cur.execute("SELECT email, active FROM users WHERE email IN (?, ?)",
            ("${TEST_ADMIN_EMAIL}", "${TEST_USER_EMAIL}"))
rows = cur.fetchall()
print(f"Seeded users: {rows}")
conn.close()
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
