// SPDX-License-Identifier: Apache-2.0
//
// vendor-runtime.mjs - fetch the same-origin runtime assets from a
// pyspark-connect-web release into public/ (git-ignored). Run: npm run vendor:runtime
//
// What this does reliably (from the release tarball):
//   - public/vendor/worker_bootstrap.js, bridge.js, coi-serviceworker.js  (glue)
//   - public/pyspark_connect_web-<ver>-py3-none-any.whl                   (pcw wheel)
//
// What you must still place same-origin (NOT in the release tarball), version-
// matched to the release (this script prints the exact target paths):
//   - public/pyodide/            (Pyodide distribution incl. pyodide.mjs)
//   - public/pyspark_client-<ver>-py3-none-any.whl
//
// See docs/reuse.md. A cross-origin CDN does NOT work for these under COI.

import { execSync } from "node:child_process";
import { mkdirSync, copyFileSync, readdirSync, rmSync, existsSync, writeFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, "..");
const PUBLIC = resolve(ROOT, "public");
const VENDOR = resolve(PUBLIC, "vendor");

const REPO = "HyukjinKwon/pyspark-client-wasm";
const TAG = process.env.PCW_TAG || "v0.1.0";
const VER = TAG.replace(/^v/, "");
const SITE_ASSET = `pyspark-connect-web-site-${VER}.tgz`;
const SITE_URL = `https://github.com/${REPO}/releases/download/${TAG}/${SITE_ASSET}`;

const tmp = resolve(ROOT, ".vendor-tmp");
rmSync(tmp, { recursive: true, force: true });
mkdirSync(tmp, { recursive: true });

console.log(`[vendor] downloading ${SITE_URL}`);
const res = await fetch(SITE_URL);
if (!res.ok) {
  console.error(`[vendor] download failed: HTTP ${res.status}. Check the tag (PCW_TAG=${TAG}).`);
  process.exit(1);
}
const tgz = resolve(tmp, SITE_ASSET);
writeFileSync(tgz, Buffer.from(await res.arrayBuffer()));

console.log("[vendor] extracting...");
execSync(`tar xzf ${JSON.stringify(tgz)} -C ${JSON.stringify(tmp)}`, { stdio: "inherit" });

// Glue -> public/vendor/
mkdirSync(VENDOR, { recursive: true });
for (const [from, to] of [
  ["worker/worker_bootstrap.js", "worker_bootstrap.js"],
  ["worker/bridge.js", "bridge.js"],
  ["coi-serviceworker.js", "coi-serviceworker.js"],
]) {
  copyFileSync(resolve(tmp, from), resolve(VENDOR, to));
  console.log(`[vendor] glue -> public/vendor/${to}`);
}

// pcw wheel -> public/
const wheel = readdirSync(tmp).find((f) => /^pyspark_connect_web-.*\.whl$/.test(f));
if (wheel) {
  copyFileSync(resolve(tmp, wheel), resolve(PUBLIC, wheel));
  console.log(`[vendor] pcw wheel -> public/${wheel}`);
} else {
  console.warn("[vendor] WARNING: pcw wheel not found in release tarball");
}

rmSync(tmp, { recursive: true, force: true });

const havePyodide = existsSync(resolve(PUBLIC, "pyodide", "pyodide.mjs"));
console.log("\n[vendor] Done with the release assets. Remaining (same-origin, version-matched):");
console.log(
  `  1. Pyodide  -> public/pyodide/ (incl. pyodide.mjs) ${havePyodide ? "[present]" : "[MISSING]"}\n` +
    "     Download the Pyodide release that pyspark-connect-web v" +
    VER +
    " targets\n" +
    "     (see its build), extract the 'full' dist into public/pyodide/.\n" +
    "  2. pyspark-client wheel -> public/pyspark_client-4.1.2-py3-none-any.whl\n" +
    "     Try: pip download pyspark-client==4.1.2 --no-deps -d public/\n" +
    "     (rename to the bootstrap's expected name if needed).\n" +
    "\n  Then: npm run build  (public/ is copied into dist/).",
);
