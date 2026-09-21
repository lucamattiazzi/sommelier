import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { once } from "node:events";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { createServer } from "node:net";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { promisify } from "node:util";
import { chromium, expect } from "@playwright/test";
import { createPairServer } from "../apps/server/dist/index.cjs";

const run = promisify(execFile);
test("task pane allows Excel frame ancestors and rejects unrelated sites", async (t) => {
  const directory = await mkdtemp(join(tmpdir(), "sommelier-framing-"));
  await writeFile(join(directory, "taskpane.html"), "<h1>Framing fixture</h1>");
  const server = await createPairServer({
    port: 0,
    publicOrigin: "https://addin.example.test",
    staticDirectory: directory,
  });
  const browser = await chromium.launch();
  const cases = [
    ["https://onedrive.live.com", "https://excel.officeapps.live.com", true],
    ["https://excel.cloud.microsoft", "https://excel.officeapps.live.com", true],
    ["https://www.microsoft365.com", "https://excel.officeapps.live.com", true],
    ["https://tenant.sharepoint.com", "https://excel.officeapps.live.com", true],
    ["https://www.office.com", "https://excel.officeapps.live.com", true],
    ["https://untrusted.example.test", "https://excel.officeapps.live.com", false],
    ["https://onedrive.live.com", "https://untrusted.example.test", false],
    ["https://office.com.untrusted.example.test", "https://excel.officeapps.live.com", false],
  ];
  try {
    for (const [outer, inner, allowed] of cases) {
      await t.test(`${outer} → ${inner}: ${allowed ? "allowed" : "blocked"}`, async () => {
        const page = await browser.newPage();
        try {
          await page.route(`${outer}/host`, (route) =>
            route.fulfill({
              contentType: "text/html",
              body: `<iframe src="${inner}/excel"></iframe>`,
            }),
          );
          await page.route(`${inner}/excel`, (route) =>
            route.fulfill({
              contentType: "text/html",
              body: '<iframe src="https://addin.example.test/taskpane.html"></iframe>',
            }),
          );
          await page.route("https://addin.example.test/taskpane.html", async (route) =>
            route.fulfill({
              response: await page.request.get(
                `http://${server.host}:${server.port}/taskpane.html`,
              ),
            }),
          );
          const blocked = allowed
            ? undefined
            : page.waitForEvent("console", (message) => message.text().includes("frame-ancestors"));
          await page.goto(`${outer}/host`);
          if (allowed) {
            await expect(
              page.frameLocator("iframe").frameLocator("iframe").getByRole("heading"),
            ).toHaveText("Framing fixture");
          } else {
            await blocked;
            for (const frame of page.frames()) {
              assert.equal(
                await frame.getByRole("heading", { name: "Framing fixture" }).count(),
                0,
              );
            }
          }
        } finally {
          await page.close();
        }
      });
    }
  } finally {
    await browser.close();
    await server.close();
    await rm(directory, { recursive: true, force: true });
  }
});

test(
  "task pane pairs, approves encrypted workbook changes and remembers terminals",
  { timeout: 60_000 },
  async () => {
    const probe = createServer();
    probe.listen(0, "127.0.0.1");
    await once(probe, "listening");
    const port = probe.address().port;
    await new Promise((done) => probe.close(done));
    const origin = `http://127.0.0.1:${port}`;
    const server = await createPairServer({
      port,
      publicOrigin: origin,
      staticDirectory: resolve("apps/addin/dist"),
    });
    const browser = await chromium.launch();
    const profile = await mkdtemp(join(tmpdir(), "pair-browser-"));
    const environment = { ...process.env, SOMMELIER_HOME: profile };
    delete environment.SOMMELIER_URL;
    delete environment.SOMMELIER_SESSION;
    const bridge = (args, url) =>
      run(
        process.execPath,
        [process.env.SOMMELIER_TEST_BRIDGE ?? "skills/sommelier/scripts/session.mjs", ...args],
        {
          env: { ...environment, ...(url ? { SOMMELIER_URL: url } : {}) },
        },
      ).then(({ stdout }) => JSON.parse(stdout));
    const rpc = (method, params) =>
      bridge(["request", "--name", "desk", "--method", method, "--params", JSON.stringify(params)]);
    const page = await browser.newPage({ viewport: { width: 320, height: 640 } });
    await page.addInitScript(() => {
      Object.defineProperty(navigator, "clipboard", {
        value: {
          writeText: async (text) => {
            window.copiedSetupPrompt = text;
          },
        },
      });
    });
    const errors = [];
    page.on("pageerror", (error) => errors.push(error.message));
    // A browser preview deliberately uses the in-memory adapter, without the Office bootstrap.
    await page.route("https://appsforoffice.microsoft.com/**", (route) => route.abort());
    try {
      await page.goto(origin);
      await expect(
        page.getByRole("heading", { name: "Work in Excel with YOUR personal agent." }),
      ).toBeVisible();
      await expect(page.getByRole("link", { name: "Download Excel add-in" })).toHaveAttribute(
        "href",
        "/manifest.xml",
      );
      assert.equal(await page.locator('script[src*="office.js"]').count(), 0);
      const manifestResponse = await fetch(`${origin}/manifest.xml`);
      assert.equal(manifestResponse.status, 200);
      const manifest = await manifestResponse.text();
      assert.ok(manifest.includes('/taskpane.html"'));
      assert.ok(!manifest.includes("localhost"));
      const linkedAssets = [...manifest.matchAll(/DefaultValue="(https:[^"]+)"/g)];
      for (const [, url] of linkedAssets) {
        const asset = await fetch(`${origin}${new URL(url).pathname}`);
        assert.equal(asset.status, 200, url);
        const head = await fetch(`${origin}${new URL(url).pathname}`, { method: "HEAD" });
        assert.equal(head.status, 200, url);
      }
      await page.setViewportSize({ width: 1100, height: 850 });
      await page.screenshot({ path: "artifacts/sommelier-home-desktop.png", fullPage: true });
      await page.emulateMedia({ colorScheme: "dark" });
      await page.screenshot({ path: "artifacts/sommelier-home-dark.png", fullPage: true });
      await page.emulateMedia({ colorScheme: "light" });
      await page.setViewportSize({ width: 320, height: 640 });
      await page.screenshot({ path: "artifacts/sommelier-home-320.png", fullPage: true });
      assert.equal(
        await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth),
        true,
      );
      await expect(page.getByRole("link", { name: "Self-host", exact: true })).toHaveAttribute(
        "href",
        "/self-host.html",
      );
      for (const path of ["/setup.html", "/privacy.html", "/support.html", "/self-host.html"]) {
        await page.goto(`${origin}${path}`);
        await expect(page.getByRole("heading", { level: 1 })).toBeVisible();
        await expect(page.getByRole("link", { name: "Sommelier", exact: true })).toHaveAttribute(
          "href",
          "/",
        );
        await expect(page.locator('nav [aria-current="page"]')).toHaveAttribute("href", path);
        assert.equal(await page.locator("script").count(), 0);
        for (const theme of ["light", "dark"]) {
          await page.emulateMedia({ colorScheme: theme });
          await expect(page.locator("html")).toHaveCSS(
            "background-color",
            theme === "light" ? "rgb(255, 255, 255)" : "rgb(22, 24, 29)",
          );
          assert.equal(
            await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth),
            true,
          );
          await page.screenshot({
            path: `artifacts/sommelier-${path.slice(1, -5)}-${theme}-320.png`,
            fullPage: true,
          });
        }
      }
      await page.emulateMedia({ colorScheme: "light" });
      await page.goto(`${origin}/taskpane.html`);
      await expect(page.getByRole("heading", { name: "Connect your agent" })).toBeVisible();
      await mkdir("artifacts/pair-review", { recursive: true });
      await page.screenshot({ path: "artifacts/pair-review/initial-320.png", fullPage: true });
      await expect(page.getByRole("textbox")).toHaveCount(0);
      await page.getByRole("button", { name: "Copy agent prompt", exact: true }).click();
      await expect(
        page.getByText("Prompt copied. Paste it into your agent.", { exact: true }),
      ).toBeVisible();
      await expect(
        page.getByRole("button", { name: "Copy agent prompt", exact: true }),
      ).toBeVisible();
      await page.getByText("Show prompt", { exact: true }).click();
      const prompt = await page.getByLabel("Prompt for your agent", { exact: true }).inputValue();
      assert.ok(prompt.includes("@lucamattiazzi/sommelier@0.2.0-beta.1"));
      assert.equal(await page.evaluate(() => window.copiedSetupPrompt), prompt);
      assert.ok(!prompt.includes("SKILL.md"));
      await page.screenshot({ path: "artifacts/pair-review/pairing-320.png", fullPage: true });
      const url = prompt.match(/wss?:\/\/[^\s]+/)[0];
      await page.getByText("Show prompt", { exact: true }).click();
      await bridge(["start", "--name", "desk"], url);
      await expect(
        page.getByText("Connected · end-to-end encrypted", { exact: true }),
      ).toBeVisible();
      await page
        .getByRole("textbox", { name: "Message your paired agent" })
        .fill("Read my synthetic selection");
      await page.getByRole("button", { name: "Send", exact: true }).click();
      const next = await bridge(["next", "--name", "desk", "--timeout", "1000"]);
      assert.equal(next.message.content, "Read my synthetic selection");
      const range = { sheetId: "Sheet1", address: "B2" };
      assert.deepEqual((await rpc("excel.range.read", { range })).result.values, [[10]]);
      const preview = await rpc("excel.operation.preview", {
        method: "excel.range.write",
        params: { range, values: [[42]] },
      });
      assert.equal(preview.ok, true);
      const committed = rpc("excel.operation.commit", {
        operationId: preview.result.operationId,
      }).catch((error) => ({ error: error.message }));
      await expect(page.getByRole("alertdialog")).toBeVisible();
      await expect(page.getByRole("button", { name: "Reject", exact: true })).toBeFocused();
      await page.keyboard.press("Shift+Tab");
      await expect(page.getByRole("button", { name: "Approve once", exact: true })).toBeFocused();
      await page.screenshot({ path: "artifacts/pair-review/approval-320.png", fullPage: true });
      await page.getByRole("button", { name: "Approve once", exact: true }).click();
      assert.equal((await committed).result.status, "committed");
      assert.deepEqual((await rpc("excel.range.read", { range })).result.values, [[42]]);
      const chartPreview = await rpc("excel.operation.preview", {
        method: "excel.chart.create",
        params: {
          range: { sheetId: "Sheet1", address: "A1:B3" },
          name: "Sales",
          title: "Synthetic sales",
          chartType: "column",
        },
      });
      const chartCommit = rpc("excel.operation.commit", {
        operationId: chartPreview.result.operationId,
      });
      await expect(page.getByRole("alertdialog")).toContainText('Create column chart "Sales"');
      await page.screenshot({
        path: "artifacts/pair-review/chart-approval-320.png",
        fullPage: true,
      });
      await page.getByRole("button", { name: "Approve once", exact: true }).click();
      assert.equal((await chartCommit).result.status, "committed");
      assert.equal(
        (await rpc("excel.chart.list", { sheetId: "Sheet1" })).result.charts[0].name,
        "Sales",
      );
      await page.getByText("Options", { exact: true }).click();
      await page.getByText("Recent operations (2)", { exact: true }).click();
      await expect(page.getByRole("button", { name: "Undo", exact: true })).toHaveCount(1);
      await page.getByText("Options", { exact: true }).click();
      await bridge(["reply", "--name", "desk", "--content", "Updated the synthetic amount to 42."]);
      await expect(
        page.getByText("Updated the synthetic amount to 42.", { exact: true }),
      ).toBeVisible();
      for (let i = 0; i < 8; i++) {
        await bridge([
          "reply",
          "--name",
          "desk",
          "--content",
          `Synthetic message ${i}. `.repeat(20),
        ]);
      }
      await expect(page.getByText(/Synthetic message 7/)).toBeInViewport();
      await expect(page.getByRole("button", { name: "Send", exact: true })).toBeInViewport();
      assert.equal(
        await page.evaluate(() => document.documentElement.scrollHeight <= innerHeight),
        true,
      );
      await page.screenshot({ path: "artifacts/pair-review/connected-320.png", fullPage: true });
      assert.equal(
        await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth),
        true,
      );
      const saved = await page.evaluate(() =>
        JSON.parse(localStorage.getItem("ai-cdl-pair-terminals-v1")),
      );
      assert.equal(saved.length, 1);
      assert.equal(saved[0].name, "My terminal");
      await page.getByText("Options", { exact: true }).click();
      await page.getByLabel("Connection name", { exact: true }).fill("Codex · Synthetic desk");
      await page.getByLabel("Connection name", { exact: true }).press("Tab");
      await page.reload();
      await page.getByText("Options", { exact: true }).click();
      await expect(page.getByText("Codex · Synthetic desk", { exact: true })).toBeVisible();
      await expect(page.getByRole("button", { name: "Reconnect", exact: true })).toBeVisible();
      await expect(page.getByRole("button", { name: "Copy connection URL" })).toHaveCount(0);
      await page.getByRole("button", { name: "Reconnect", exact: true }).click();
      await expect(
        page.getByText("Connected · end-to-end encrypted", { exact: true }),
      ).toBeVisible();
      await page.getByText("Options", { exact: true }).click();
      await page.getByText("Manage connection", { exact: true }).click();
      await page.getByLabel("Connect automatically when this pane opens").check();
      await page
        .getByLabel("Approve changes automatically for this connection only. Previews still run.")
        .check();
      await page.reload();
      await expect(
        page.getByText("Connected · end-to-end encrypted", { exact: true }),
      ).toBeVisible();
      await expect(
        page.getByLabel(
          "Approve changes automatically for this connection only. Previews still run.",
        ),
      ).not.toBeChecked();
      await page.getByText("Options", { exact: true }).click();
      await page.getByText("Manage connection", { exact: true }).click();
      await page.getByRole("button", { name: "Forget", exact: true }).click();
      await expect(page.getByRole("button", { name: "Reconnect", exact: true })).toHaveCount(0);
      assert.equal(
        await page.evaluate(
          () => JSON.parse(localStorage.getItem("ai-cdl-pair-terminals-v1")).length,
        ),
        0,
      );
      await page.evaluate(() => {
        navigator.clipboard.writeText = async () => {
          throw new Error("Clipboard blocked");
        };
      });
      await page.getByRole("button", { name: "Copy agent prompt", exact: true }).click();
      await expect(page.getByRole("alert")).toContainText("Select and copy the prompt below");
      await expect(page.getByLabel("Prompt for your agent", { exact: true })).toBeVisible();
      assert.deepEqual(errors, []);
      for (const path of [
        "/taskpane.html",
        "/landing.css",
        "/setup.html",
        "/privacy.html",
        "/support.html",
        "/self-host.html",
        "/help.css",
        "/icon-16.png",
        "/icon-32.png",
        "/icon-64.png",
        "/icon-80.png",
        "/icon-128.png",
        "/manifest.xml",
        "/agent/session.mjs",
        "/agent/lib/encrypted-socket.mjs",
      ]) {
        const response = await fetch(`${origin}${path}`);
        assert.equal(response.status, 200, path);
        assert.ok((await response.arrayBuffer()).byteLength > 0, path);
      }
    } finally {
      await bridge(["stop", "--name", "desk"]).catch(() => undefined);
      await browser.close();
      await server.close();
      await rm(profile, { recursive: true, force: true });
    }
  },
);
