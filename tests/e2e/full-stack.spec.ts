// SPDX-License-Identifier: Apache-2.0
//
// full-stack.spec.ts - the COMPLETE browser-to-Spark e2e.
//
// Drives the real web demo in headless Chromium: boots Pyodide +
// pyspark-connect-web, connects through the Envoy grpc-web proxy to a real
// Spark Connect server, runs a Spark SQL query, and asserts the rendered
// result. This exercises the entire stack that the COI gate and the Python
// integration test each cover only part of.
//
// Runs only when E2E_FULL=1 (the deploy stack must be up + dist vendored with
// Pyodide and the wheels). The e2e-full.yml workflow sets that up.

import { test, expect } from "@playwright/test";

const FULL = process.env.E2E_FULL === "1";
const DEMO_URL = process.env.E2E_DEMO_URL || "http://localhost:8000/demo/demo.html";

test.describe("full-stack: Pyodide + grpc-web + Spark Connect", () => {
  test.skip(!FULL, "set E2E_FULL=1 with the deploy stack up + assets vendored");

  test("connect and run a Spark SQL query, end to end", async ({ page }) => {
    test.setTimeout(780_000); // Pyodide 314 cold boot + wheel install + first Spark call is slow

    // Capture the full browser console + page errors so a failure shows WHY
    // (Pyodide boot / wheel install / grpc-web errors land here).
    const log: string[] = [];
    page.on("console", (m) => log.push(`[page:${m.type()}] ${m.text()}`));
    page.on("pageerror", (e) => log.push(`[pageerror] ${e.message}`));
    // Pyodide + pyspark-connect-web run in a Web Worker - capture ITS console too
    // (the boot/connect output lives there, not on the main thread).
    page.on("worker", (w) => {
      w.on("console", (m) => log.push(`[worker:${m.type()}] ${m.text()}`));
    });

    await page.goto(DEMO_URL);

    // Served with COOP/COEP -> cross-origin isolated (prereq for SharedArrayBuffer).
    await expect
      .poll(() => page.evaluate(() => self.crossOriginIsolated), { timeout: 20_000 })
      .toBe(true);

    // Connection form -> Envoy grpc-web endpoint.
    await page.fill("input.demo-input-host", "localhost");
    await page.fill("input.demo-input-port", "8081");
    await page.getByRole("button", { name: "Connect" }).click();

    // Booting Pyodide + installing wheels + connecting is slow on a cold runner.
    const connStatus = page.locator(".demo-status").first();
    try {
      await expect(connStatus).toContainText("Connected", { timeout: 660_000 });
    } catch (e) {
      const txt = await connStatus.textContent().catch(() => "(no status)");
      console.log("=== CONNECT did not complete. Last status:", txt);
      console.log("=== Browser log:\n" + log.join("\n"));
      await page.screenshot({ path: "playwright-report/connect-failed.png", fullPage: true });
      throw e;
    }

    // Run a real Spark SQL query that yields a multi-row, chartable result so
    // the captured screenshot shows an actual rendered chart (category + value).
    const sql =
      "SELECT CONCAT('Q', CAST(id + 1 AS STRING)) AS quarter, " +
      "CAST((id + 1) * (id + 1) * 1000 AS INT) AS revenue FROM range(0, 6)";
    await page.locator(".code-editor__textarea").fill(sql);
    await page.getByRole("button", { name: "Run" }).click();

    // Result table rendered from real Spark output.
    const table = page.locator("table.demo-table");
    await expect(table).toBeVisible({ timeout: 60_000 });
    await expect(table.locator("thead th").nth(0)).toContainText("quarter");
    await expect(table.locator("thead th").nth(1)).toContainText("revenue");
    await expect(table.locator("tbody tr").nth(0).locator("td").nth(0)).toContainText("Q1");
    // The demo renders a real SVG chart from the result (category + numeric).
    await expect(page.locator(".demo-chart svg")).toBeVisible({ timeout: 30_000 });

    // Capture a genuine screenshot of the running demo (uploaded as an artifact):
    // a live Spark SQL query with the rendered chart.
    await page.screenshot({ path: "playwright-report/web-demo-live.png", fullPage: true });
  });
});
