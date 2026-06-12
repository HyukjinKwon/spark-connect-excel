// SPDX-License-Identifier: Apache-2.0
//
// render-icons.mjs - rasterize public/assets/icon.svg into the PNG sizes the
// Office manifest references. Run: npm run icons
//
// Uses the Playwright Chromium that the e2e already installs, so there is no new
// dependency. Transparent background (omitBackground) so the spark sits cleanly
// on the Excel ribbon and the Insert > Add-ins dialog.

import { readFileSync, writeFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ASSETS = resolve(__dirname, "..", "public", "assets");
const SIZES = [16, 32, 64, 80, 128];

const svg = readFileSync(resolve(ASSETS, "icon.svg"), "utf8");

// Allow pointing at an already-installed Chromium (PLAYWRIGHT_CHROMIUM) so this
// works without re-downloading the exact browser build the package pins.
const execPath = process.env.PLAYWRIGHT_CHROMIUM || undefined;
const browser = await chromium.launch(execPath ? { executablePath: execPath } : {});
try {
  for (const size of SIZES) {
    const page = await browser.newPage({ viewport: { width: size, height: size } });
    const sized = svg.replace(
      /<svg([^>]*?)width="\d+"\s+height="\d+"/,
      `<svg$1width="${size}" height="${size}"`,
    );
    await page.setContent(
      `<!doctype html><html><head><style>*{margin:0;padding:0}html,body{background:transparent}svg{display:block}</style></head><body>${sized}</body></html>`,
      { waitUntil: "networkidle" },
    );
    const buf = await page.locator("svg").screenshot({ omitBackground: true });
    writeFileSync(resolve(ASSETS, `icon-${size}.png`), buf);
    console.log(`[icons] wrote public/assets/icon-${size}.png`);
    await page.close();
  }
} finally {
  await browser.close();
}
