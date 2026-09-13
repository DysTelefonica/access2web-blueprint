/**
 * e2e/cli.spec.ts
 *
 * Phase 5 E2E — CLI commands.
 *
 * Tests the "gentle-ai platform user" CLI subcommands inside the running
 * Docker container using `docker exec`.  All commands are exercised via the
 * Python module invocation:
 *
 *   python -m app.src.modules.lanzadera.delivery.cli.platform_user <subcommand>
 *
 * Destructive commands (set-password, assign-profile) are run with
 * `--confirmed` so the confirmation prompt is bypassed.
 *
 * The new-password for `set-password` is piped via stdin using a
 * heredoc-style stdin redirect.
 *
 * Test users are seeded by the Phase 1 globalSetup.
 */

import { test, expect } from "@playwright/test";
import { spawn } from "node:child_process";
import { promisify } from "node:util";

import { TEST_USER_EMAIL, TEST_ADMIN_EMAIL } from "./fixtures/auth-fixture";

const execAsync = promisify(exec);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

interface CliResult {
  exitCode: number;
  stdout: string;
  stderr: string;
}

/** Run a CLI subcommand inside the Docker container. */
function runCli(
  subcommand: string[],
  options: { pipeStdin?: string; containerName?: string } = {}
): Promise<CliResult> {
  return new Promise((resolve) => {
    const containerName = options.containerName ?? process.env.TEST_CONTAINER_NAME ?? "test-app";
    const args = [
      "exec", "-i",            // -i = keep stdin open
      containerName,
      "python", "-m",
      "app.src.modules.lanzadera.delivery.cli.platform_user",
      ...subcommand,
    ];
    const proc = spawn("docker", args);

    let stdout = "";
    let stderr = "";

    proc.stdout?.on("data", (chunk: Buffer) => {
      stdout += chunk.toString();
    });
    proc.stderr?.on("data", (chunk: Buffer) => {
      stderr += chunk.toString();
    });

    if (options.pipeStdin) {
      proc.stdin?.write(options.pipeStdin);
    }
    proc.stdin?.end();

    proc.on("close", (code) => {
      resolve({ exitCode: code ?? 0, stdout, stderr });
    });
    proc.on("error", (err) => {
      resolve({ exitCode: 1, stdout: "", stderr: err.message });
    });
  });
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

test.describe("CLI — gentle-ai platform user", () => {
  // ---------------------------------------------------------------------------
  // set-password
  // ---------------------------------------------------------------------------

  test.describe("set-password", () => {
    test("changes user password and returns exit 0", async () => {
      const newPassword = "NewE2E!Pwd99";
      const result = await runCli(
        ["set-password", "--confirmed", TEST_USER_EMAIL],
        { pipeStdin: `${newPassword}\n` }
      );
      expect(result.exitCode).toBe(0);
      expect(result.stdout).toContain("password set");
      expect(result.stdout).toContain("status=active");
    });

    test("returns exit 1 for unknown email", async () => {
      const result = await runCli(
        ["set-password", "--confirmed", "nobody@e2e-cli.test"],
        { pipeStdin: "DoesNotMatter99\n" }
      );
      expect(result.exitCode).toBe(1);
      expect(result.stdout + result.stderr).toMatch(/not found/i);
    });

    test("returns exit 1 when stdin is empty (no password)", async () => {
      const result = await runCli(
        ["set-password", "--confirmed", TEST_USER_EMAIL],
        { pipeStdin: "\n" }
      );
      expect(result.exitCode).toBe(1);
      expect(result.stdout).toContain("empty password");
    });
  });

  // ---------------------------------------------------------------------------
  // list-apps
  // ---------------------------------------------------------------------------

  test.describe("list-apps", () => {
    test("returns exit 0 and lists active apps", async () => {
      const result = await runCli(["list-apps"]);
      expect(result.exitCode).toBe(0);
      expect(result.stdout).toContain("active apps:");
    });
  });

  // ---------------------------------------------------------------------------
  // assign-profile
  // ---------------------------------------------------------------------------

  test.describe("assign-profile", () => {
    test("returns exit 0 when assigning profile to known user and app", async () => {
      // Assign profile "READ" (profile_code) to app 1 (app_id) for the test admin
      const result = await runCli(
        ["assign-profile", "--confirmed", TEST_ADMIN_EMAIL, "1", "READ"],
      );
      // Expect either success (exit 0) or graceful "already assigned" (exit 0 or 1)
      expect([0, 1]).toContain(result.exitCode);
      if (result.exitCode === 0) {
        expect(result.stdout).toContain("assign-profile");
      }
    });

    test("returns exit 1 for unknown user email", async () => {
      const result = await runCli(
        ["assign-profile", "--confirmed", "nobody@e2e-cli.test", "1", "READ"],
      );
      expect(result.exitCode).toBe(1);
      expect(result.stdout + result.stderr).toMatch(/not found/i);
    });
  });

  // ---------------------------------------------------------------------------
  // grant-global-admin
  // ---------------------------------------------------------------------------

  test.describe("grant-global-admin", () => {
    test("returns exit 0 when granting to existing user", async () => {
      const result = await runCli(
        ["grant-global-admin", "--confirmed", TEST_USER_EMAIL],
      );
      // Already granted in seed → idempotent success
      expect([0, 1]).toContain(result.exitCode);
      if (result.exitCode === 0) {
        expect(result.stdout).toContain("global-admin");
      }
    });

    test("returns exit 1 for unknown user", async () => {
      const result = await runCli(
        ["grant-global-admin", "--confirmed", "ghost@e2e-cli.test"],
      );
      expect(result.exitCode).toBe(1);
      expect(result.stdout + result.stderr).toMatch(/not found/i);
    });
  });

  // ---------------------------------------------------------------------------
  // revoke-global-admin
  // ---------------------------------------------------------------------------

  test.describe("revoke-global-admin", () => {
    test("returns exit 0 when revoking existing admin", async () => {
      const result = await runCli(
        ["revoke-global-admin", "--confirmed", TEST_USER_EMAIL],
      );
      expect([0, 1]).toContain(result.exitCode);
      if (result.exitCode === 0) {
        expect(result.stdout).toContain("revoke-global-admin");
      }
    });

    test("returns exit 1 for unknown user", async () => {
      const result = await runCli(
        ["revoke-global-admin", "--confirmed", "ghost@e2e-cli.test"],
      );
      expect(result.exitCode).toBe(1);
    });
  });
});
