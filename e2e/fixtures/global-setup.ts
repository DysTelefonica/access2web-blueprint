/**
 * e2e/fixtures/global-setup.ts
 *
 * Playwright 1.64+ requires the file referenced by `globalSetup` to export a
 * SINGLE function via `export default`. `auth-fixture.ts` exports several
 * helpers (TEST_ADMIN_*, getAdminToken, getUserToken) that other specs
 * import directly, so it cannot satisfy the "single export" rule without
 * losing its public surface.
 *
 * This module is the single-export-shape sibling: it imports only the
 * credentials it needs, keeps the `runInContainer` helper private, and
 * re-exports nothing else.
 *
 * Responsibilities:
 *  1. Run Alembic migrations against the live container before any test fires.
 *     This ensures the DB schema is up-to-date regardless of how the image was
 *     built (the Dockerfile does not bake in migrations).
 *  2. Seed known admin and regular users directly into SQLite.
 *     The seed script runs Python inside the container (where argon2-cffi is
 *     available) so it can hash passwords with the real hasher.
 *
 * Runtime contract (read from env):
 *  - `TEST_CONTAINER_NAME` — defaults to `test-app` (matches the e2e-smoke
 *    job and the release workflow).
 */

import { type FullConfig } from "@playwright/test";

import {
  TEST_ADMIN_EMAIL,
  TEST_ADMIN_NAME,
  TEST_ADMIN_PASSWORD,
  TEST_USER_EMAIL,
  TEST_USER_NAME,
  TEST_USER_PASSWORD,
} from "./auth-fixture";

// ---------------------------------------------------------------------------
// Docker exec helper — kept private to this module (only globalSetup uses it)
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

// ---------------------------------------------------------------------------
// globalSetup — Playwright runs this once per worker before any test
// ---------------------------------------------------------------------------

export default async function globalSetup(_config: FullConfig): Promise<void> {
  const containerName = process.env.TEST_CONTAINER_NAME ?? "test-app";

  // 1. Apply schema migrations
  await runInContainer(containerName, [
    "python", "-m", "alembic", "upgrade", "head",
  ]);
  console.log("[e2e:globalSetup] Migrations applied.");

  // 2. Seed test users — Python script runs inside the container where
  //    argon2-cffi is available, so we get real hashes for the test passwords.
  const seedScript = `
import sqlite3, sys, uuid
from argon2 import PasswordHash

# Generate real Argon2id hashes using the container's argon2-cffi
h = PasswordHash()
admin_hash = h.hash("${TEST_ADMIN_PASSWORD}")
user_hash  = h.hash("${TEST_USER_PASSWORD}")

conn = sqlite3.connect("test.db")
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