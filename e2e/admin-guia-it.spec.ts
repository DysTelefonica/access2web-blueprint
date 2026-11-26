/**
 * e2e/admin-guia-it.spec.ts
 *
 * Admin IT handover guide (Citrix + Linux), #1810.
 *
 * Tests cover:
 *  - Reaching /admin/guia-it-citrix-linux from the admin navigation
 *  - Following a glossary link to its definition
 *  - The guide is no-store and closed to non-admin users
 */

import { test, expect } from "@playwright/test";

import { getAdminToken, getUserToken } from "./fixtures/auth-fixture";

const GUIDE = "/admin/guia-it-citrix-linux";

test.describe("Admin IT handover guide", () => {
  let adminToken: string;
  let userToken: string;

  test.beforeAll(async () => {
    adminToken = await getAdminToken();
    userToken = await getUserToken();
  });

  test("an admin reaches the guide from the navigation and follows a glossary link", async ({
    page,
  }) => {
    await page.setExtraHTTPHeaders({ Authorization: `Bearer ${adminToken}` });
    await page.goto("/admin/audit");
    await page.getByRole("link", { name: "Guía IT" }).click();

    await expect(page).toHaveURL(new RegExp(`${GUIDE}$`));
    await expect(
      page.getByRole("heading", { name: "Citrix, Linux y coexistencia" })
    ).toBeVisible();

    await page.locator('a.guia-it__term[href="#g-macros"]').first().click();
    await expect(page).toHaveURL(/#g-macros$/);
    await expect(page.locator("#g-macros")).toBeVisible();
  });

  test("the guide is not cached and is closed to non-admin users", async ({
    request,
  }) => {
    const admin = await request.get(GUIDE, {
      headers: { Authorization: `Bearer ${adminToken}` },
    });
    expect(admin.status()).toBe(200);
    expect(admin.headers()["cache-control"]).toBe("no-store");

    const user = await request.get(GUIDE, {
      headers: { Authorization: `Bearer ${userToken}` },
    });
    expect(user.status()).toBe(403);

    const anonymous = await request.get(GUIDE);
    expect(anonymous.status()).toBe(401);
  });
});
