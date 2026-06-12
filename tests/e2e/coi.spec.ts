// SPDX-License-Identifier: Apache-2.0
//
// coi.spec.ts — LIVE cross-origin isolation gate.
//
// This is the make-or-break check for the whole spark-connect-excel
// architecture. The pyspark-connect-web SAB bridge (the backbone of
// blocking .collect()) requires SharedArrayBuffer, which is only available in
// cross-origin-isolated contexts:
//
//   Cross-Origin-Opener-Policy: same-origin
//   Cross-Origin-Embedder-Policy: credentialless   (DECISIONS #2)
//
// These tests navigate to the dialog page (/dialog/dialog.html) — the same
// page that Excel opens as the COI host window — and assert:
//
//   1. self.crossOriginIsolated === true
//   2. typeof SharedArrayBuffer !== "undefined"
//
// Both checks can run WITHOUT Excel or a running Spark server. They rely only
// on `vite preview` serving dist/ with the correct headers (from vite.config.ts).
//
// If either assertion fails:
//   - Check that `npm run build` ran before `playwright test`.
//   - Check that vite.config.ts still sets COOP/COEP headers in the `preview`
//     block (the CI COI-header-guard step catches regressions here).
//   - Check that the dialog page loads at all (no JS parse errors before the
//     check runs — dialog.html loads Office.js from the CDN which is unavailable
//     in a plain Playwright context; the page is expected to partially fail, but
//     the COI/SAB check must still pass since it doesn't depend on Office).
//
// Office.js will likely fail to load (CDN unreachable or no Office host), but
// that happens AFTER the page is served with the correct headers, so the
// crossOriginIsolated flag is set at page load time and is readable regardless.

import { test, expect } from "@playwright/test";

const DIALOG_PAGE = "/dialog/dialog.html";

// Suppress any JavaScript errors that come from Office.js not being available
// in the Playwright context. Those errors are expected and must not block the
// COI/SAB assertion.
test.beforeEach(async ({ page }) => {
  // Register error handler BEFORE navigation so we see any errors.
  page.on("pageerror", (_err) => {
    // Intentionally ignored: Office.js CDN errors and Office.onReady failures
    // are expected when running outside an Excel host. The test only cares about
    // the COI headers, which are applied by the server before JS runs.
  });
  page.on("console", (_msg) => {
    // Suppress console noise from Office.js in test output.
  });
});

test("dialog page is cross-origin isolated (crossOriginIsolated === true)", async ({
  page,
}) => {
  // Navigate to the dialog page. `waitUntil: "domcontentloaded"` avoids waiting
  // for Office.js CDN scripts that will fail to load in this context.
  await page.goto(DIALOG_PAGE, { waitUntil: "domcontentloaded" });

  // The COI check does not depend on any script execution — it's a property
  // of the browsing context set by the response headers. Evaluate it directly.
  const crossOriginIsolated = await page.evaluate(
    () => (self as Window & typeof globalThis).crossOriginIsolated,
  );

  expect(crossOriginIsolated).toBe(true);
});

test("SharedArrayBuffer is available in the dialog page", async ({ page }) => {
  await page.goto(DIALOG_PAGE, { waitUntil: "domcontentloaded" });

  const sabAvailable = await page.evaluate(
    () => typeof SharedArrayBuffer !== "undefined",
  );

  expect(sabAvailable).toBe(true);
});

test("Atomics is available alongside SharedArrayBuffer", async ({ page }) => {
  // Atomics availability is co-gated with SAB on crossOriginIsolated pages.
  await page.goto(DIALOG_PAGE, { waitUntil: "domcontentloaded" });

  const atomicsAvailable = await page.evaluate(
    () => typeof Atomics !== "undefined",
  );

  expect(atomicsAvailable).toBe(true);
});

test("COOP header is same-origin (verified via page context)", async ({ request }) => {
  // Use the Playwright request API to check the raw response headers.
  const response = await request.get(DIALOG_PAGE);
  const coopHeader = response.headers()["cross-origin-opener-policy"];
  expect(coopHeader).toBe("same-origin");
});

test("COEP header is credentialless (verified via page context)", async ({
  request,
}) => {
  // DECISIONS #2: COEP must be 'credentialless', not 'require-corp'.
  // 'credentialless' allows Pyodide (jsDelivr) and PyPI wheels without CORP.
  const response = await request.get(DIALOG_PAGE);
  const coepHeader = response.headers()["cross-origin-embedder-policy"];
  expect(coepHeader).toBe("credentialless");
});
