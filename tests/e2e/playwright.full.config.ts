// SPDX-License-Identifier: Apache-2.0
//
// Full-stack e2e config. Unlike playwright.config.ts (which runs the COI gate
// against `vite preview`), this one has NO webServer - the deploy/ docker stack
// serves the demo at http://localhost:8000. Runs full-stack.spec.ts only.
import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: ".",
  testMatch: "full-stack.spec.ts",
  timeout: 780_000,
  expect: { timeout: 660_000 },
  retries: 0,
  workers: 1,
  reporter: [
    ["list"],
    ["html", { outputFolder: "playwright-report", open: "never" }],
  ],
  projects: [
    {
      name: "chromium",
      use: {
        ...devices["Desktop Chrome"],
        channel: process.env.PW_CHANNEL || undefined,
      },
    },
  ],
  use: { trace: "on" },
});
