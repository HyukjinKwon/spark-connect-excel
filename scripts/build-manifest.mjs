// SPDX-License-Identifier: Apache-2.0
//
// build-manifest.mjs — generate a production manifest by substituting the dev
// origin (https://localhost:3000) for your real hosting origin.
//
// Usage:
//   node scripts/build-manifest.mjs --origin https://my-addin.example.com
//   ADDIN_ORIGIN=https://my-addin.example.com node scripts/build-manifest.mjs
//   node scripts/build-manifest.mjs --origin <url> --out dist/manifest.xml
//
// With no --origin it copies the manifest unchanged (still localhost) and warns.
// Validate the result afterwards with: npx office-addin-manifest validate <out>

import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, "..");
const DEV_ORIGIN = "https://localhost:3000";

function arg(name) {
  const i = process.argv.indexOf(name);
  return i !== -1 && i + 1 < process.argv.length ? process.argv[i + 1] : undefined;
}

const origin = arg("--origin") || process.env.ADDIN_ORIGIN;
const out = arg("--out") || resolve(ROOT, "dist", "manifest.xml");
const src = resolve(ROOT, "manifest.xml");

const original = readFileSync(src, "utf8");

if (!origin) {
  console.warn(
    `[build-manifest] No --origin / ADDIN_ORIGIN given — copying manifest unchanged ` +
      `(still ${DEV_ORIGIN}). Provide a production origin before publishing.`,
  );
} else {
  if (!/^https:\/\/[^/]+$/.test(origin)) {
    console.error(
      `[build-manifest] --origin must be a bare HTTPS origin like ` +
        `https://my-addin.example.com (no trailing slash, no path). Got: ${origin}`,
    );
    process.exit(1);
  }
}

const result = origin ? original.split(DEV_ORIGIN).join(origin) : original;

mkdirSync(dirname(out), { recursive: true });
writeFileSync(out, result, "utf8");

const count = origin ? original.split(DEV_ORIGIN).length - 1 : 0;
console.log(
  `[build-manifest] Wrote ${out}` +
    (origin ? ` (replaced ${count} occurrence(s) of ${DEV_ORIGIN} → ${origin})` : ""),
);
console.log(`[build-manifest] Validate it: npx office-addin-manifest validate ${out}`);
