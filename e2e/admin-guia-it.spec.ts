// Admin IT handover guide (#1810). The seeded admin is a plain user, so the spec grants it
// global admin; Chromium forces HTTPS on the HSTS-preloaded `app` TLD (the local stack's
// host), so the page uses an alias host resolved to BASE_URL's host.

import { execFileSync } from "node:child_process";

import { test, expect } from "@playwright/test";

import { getAdminToken, getUserToken, TEST_ADMIN_EMAIL } from "./fixtures/auth-fixture";

const GUIDE = "/admin/guia-it-citrix-linux";
const TARGET = new URL(process.env.BASE_URL ?? "http://localhost:8000");
const ALIAS = "guia-it.e2e.test";
const GRANT = `import asyncio, os, sys, asyncpg
async def main():
    c = await asyncpg.connect(os.environ["DATABASE_URL"].replace("+asyncpg", "", 1))
    await c.execute("INSERT INTO lanzadera.global_admins (user_id, granted_at) SELECT id, now() FROM lanzadera.users WHERE email = $1 ON CONFLICT DO NOTHING", sys.argv[1])
asyncio.run(main())`;

test.use({ launchOptions: { args: [`--host-resolver-rules=MAP ${ALIAS} ${TARGET.hostname}`] } });

test.describe("Admin IT handover guide", () => {
  let adminToken: string;
  let userToken: string;

  test.beforeAll(async () => {
    const container = process.env.TEST_CONTAINER_NAME ?? "test-app";
    execFileSync("docker", ["exec", container, "python", "-c", GRANT, TEST_ADMIN_EMAIL]);
    adminToken = await getAdminToken();
    userToken = await getUserToken();
  });

  test("an admin reaches the guide from the navigation and follows a glossary link", async ({ page }) => {
    await page.setExtraHTTPHeaders({ Authorization: `Bearer ${adminToken}` });
    await page.goto(`http://${ALIAS}:${TARGET.port || "80"}/admin/audit`);
    await page.getByRole("link", { name: "Guía IT" }).click();

    await expect(page).toHaveURL(new RegExp(`${GUIDE}$`));
    await expect(page.getByRole("heading", { name: "Citrix, Linux y coexistencia" })).toBeVisible();

    await page.locator('a.guia-it__term[href="#g-macros"]').first().click();
    await expect(page).toHaveURL(/#g-macros$/);
    await expect(page.locator("#g-macros")).toBeVisible();
  });

  test("the guide is not cached and is closed to non-admin users", async ({ request }) => {
    const admin = await request.get(GUIDE, { headers: { Authorization: `Bearer ${adminToken}` } });
    expect(admin.status()).toBe(200);
    expect(admin.headers()["cache-control"]).toBe("no-store");

    const user = await request.get(GUIDE, { headers: { Authorization: `Bearer ${userToken}` } });
    expect(user.status()).toBe(403);

    const anonymous = await request.get(GUIDE);
    expect(anonymous.status()).toBe(401);
  });
});
