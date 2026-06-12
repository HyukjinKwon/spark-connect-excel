// SPDX-License-Identifier: Apache-2.0
//
// playwright.config.ts — Playwright configuration for the spark-connect-excel
// end-to-end test harness.
//
// What runs automatically (no Excel / Spark required):
//   - coi.spec.ts: asserts crossOriginIsolated === true and SharedArrayBuffer
//     availability against the built dialog page (vite preview serves dist/).
//
// What is deferred (requires real Excel + Spark stack):
//   - query-flow.spec.ts: guarded by the E2E_REQUIRE_STACK env var.
//
// Usage:
//   npm run build           # produce dist/ with COI headers
//   npx playwright test     # runs the COI gate (live) + deferred (skipped)
//
// See tests/e2e/README.md for the full harness documentation.

import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: ".",
  // Only pick up *.spec.ts files in this directory.
  testMatch: "**/*.spec.ts",

  // Reasonable timeout for page-level operations (not waiting for Pyodide).
  timeout: 30_000,
  // Expect assertions timeout.
  expect: { timeout: 10_000 },

  // Retry once in CI to absorb flaky port-binding races.
  retries: process.env["CI"] ? 1 : 0,

  // Run tests sequentially within a file (default); files in parallel is fine
  // since we only have two spec files and they don't share state.
  workers: 1,

  reporter: [
    ["list"],
    [
      "html",
      {
        outputFolder: "playwright-report",
        open: "never",
      },
    ],
  ],

  projects: [
    {
      name: "chromium",
      use: {
        ...devices["Desktop Chrome"],
        // Chromium is the only host that matters: Office add-ins on Windows use
        // WebView2 (Chromium-based), Excel on the web uses Chrome/Edge. The
        // COEP credentialless mode we rely on is a Chromium feature.
        //
        // PW_CHANNEL=chrome drives the system-installed Chrome (no Playwright
        // browser download needed — useful offline / in restricted networks).
        // Unset (CI) uses Playwright's bundled Chromium.
        channel: process.env["PW_CHANNEL"] || undefined,
        launchOptions: {
          // Allow SharedArrayBuffer flags if the browser needs them (not needed
          // once COI headers are set, but belt-and-suspenders for local dev).
          args: [],
        },
      },
    },
  ],

  // Web server: run `vite preview` against the pre-built dist/ directory.
  // The preview server inherits the COI headers from vite.config.ts:
  //   Cross-Origin-Opener-Policy: same-origin
  //   Cross-Origin-Embedder-Policy: credentialless
  // This is exactly what makes crossOriginIsolated === true in the dialog page.
  //
  // IMPORTANT: `npm run build` must have been run before the test suite starts.
  // CI does this in a prior step (see .github/workflows/e2e.yml).
  webServer: {
    command: "npm run preview",
    port: 3000,
    // Wait up to 30s for the preview server to start.
    timeout: 30_000,
    // Re-use an already-running preview server in local dev.
    reuseExistingServer: !process.env["CI"],
    // Kill the server after tests complete.
    stdout: "ignore",
    stderr: "pipe",
  },

  use: {
    baseURL: "http://localhost:3000",
    // No trace/video by default; enable in CI via env override if needed.
    trace: "on-first-retry",
  },
});
