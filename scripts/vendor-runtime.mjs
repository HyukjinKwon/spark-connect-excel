// SPDX-License-Identifier: Apache-2.0
//
// vendor-runtime.mjs - vendor ALL same-origin runtime assets into public/
// (git-ignored), producing the exact set the end-to-end CI gate (e2e-full.yml)
// proved working. Run: npm run vendor:runtime
//
// It mirrors the e2e steps so a local checkout matches the proven stack:
//   1. clone pyspark-connect-web (PCW_REF, default "main") - ONE checkout, so the
//      glue and the pcw wheel are version-matched by construction.
//   2. public/vendor/{worker_bootstrap.js,bridge.js,coi-serviceworker.js}  (glue)
//   3. public/pyspark_connect_web-<ver>-py3-none-any.whl   (built from the clone)
//   4. public/<the spark client wheel the bootstrap names>  (built via pip)
//   5. public/pyodide/  (Pyodide PYODIDE_VERSION, default 314.0.0, full dist)
//
// Requires: git, python3 + pip, curl, tar. A cross-origin CDN does NOT work for
// these under cross-origin isolation, which is why everything is same-origin.
//
// After this: npm run build  (public/ is copied into dist/), then bring up the
// stack with deploy/compose.yaml. See docs/reuse.md.

import { execSync } from "node:child_process";
import { mkdirSync, copyFileSync, readdirSync, rmSync, renameSync, readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, "..");
const PUBLIC = resolve(ROOT, "public");
const VENDOR = resolve(PUBLIC, "vendor");

const REPO = process.env.PCW_REPO || "https://github.com/HyukjinKwon/pyspark-client-wasm";
const REF = process.env.PCW_REF || "main";
const PYODIDE_VERSION = process.env.PYODIDE_VERSION || "314.0.0";

const sh = (cmd) => execSync(cmd, { stdio: "inherit" });
const out = (cmd) => execSync(cmd, { encoding: "utf8" }).trim();

const tmp = resolve(ROOT, ".vendor-tmp");
rmSync(tmp, { recursive: true, force: true });
mkdirSync(tmp, { recursive: true });
mkdirSync(VENDOR, { recursive: true });

// Remove any previously vendored wheels + Pyodide so a re-run never leaves a
// stale/duplicate wheel (the bootstrap names exactly one of each).
for (const f of readdirSync(PUBLIC)) {
  if (/^pyspark.*\.whl$/.test(f) || f.endsWith(".tar.gz")) rmSync(resolve(PUBLIC, f));
}
rmSync(resolve(PUBLIC, "pyodide"), { recursive: true, force: true });

const pcw = resolve(tmp, "pcw");
console.log(`[vendor] cloning ${REPO} @ ${REF}`);
sh(`git clone --depth 1 --branch ${JSON.stringify(REF)} ${JSON.stringify(REPO)} ${JSON.stringify(pcw)} 2>/dev/null || git clone --depth 1 ${JSON.stringify(REPO)} ${JSON.stringify(pcw)}`);

// 2. Glue -> public/vendor/ (from the SAME checkout as the wheel below).
for (const [from, to] of [
  ["pyspark_connect_web/worker/worker_bootstrap.js", "worker_bootstrap.js"],
  ["pyspark_connect_web/worker/bridge.js", "bridge.js"],
  ["pyspark_connect_web/jupyterlite/coi-serviceworker.js", "coi-serviceworker.js"],
]) {
  copyFileSync(resolve(pcw, from), resolve(VENDOR, to));
  console.log(`[vendor] glue -> public/vendor/${to}`);
}

// 3. pcw wheel (built from the clone -> version matches the glue).
console.log("[vendor] building the pyspark-connect-web wheel");
sh(`python3 -m pip wheel ${JSON.stringify(pcw)} --no-deps -w ${JSON.stringify(PUBLIC)}`);

// 4. The Spark client wheel the bootstrap names (pyspark-<ver> or pyspark_client-<ver>,
//    NOT pyspark_connect_web which also starts with "pyspark").
const boot = readFileSync(resolve(VENDOR, "worker_bootstrap.js"), "utf8");
const want = (boot.match(/\/pyspark(_client)?-[0-9][A-Za-z0-9._-]*\.whl/) || [])[0]?.replace(/^\//, "");
if (!want) {
  console.error("[vendor] ERROR: could not find the Spark client wheel name in worker_bootstrap.js");
  process.exit(1);
}
const isClient = want.startsWith("pyspark_client-");
const sver = want.replace(/^pyspark(_client)?-([0-9.]+)-.*/, "$2");
const pkg = isClient ? "pyspark-client" : "pyspark";
console.log(`[vendor] bootstrap wants ${want} -> building ${pkg}==${sver}`);
sh(`python3 -m pip wheel ${JSON.stringify(`${pkg}==${sver}`)} --no-deps -w ${JSON.stringify(PUBLIC)}`);
const built = readdirSync(PUBLIC).find((f) => f.startsWith(isClient ? "pyspark_client-" : "pyspark-") && f.endsWith(".whl"));
if (built && built !== want) {
  renameSync(resolve(PUBLIC, built), resolve(PUBLIC, want));
  console.log(`[vendor] renamed ${built} -> ${want} (matches the bootstrap)`);
}

// 5. Pyodide (same-origin, full dist).
console.log(`[vendor] fetching Pyodide ${PYODIDE_VERSION}`);
const pyTar = resolve(tmp, "pyodide.tar.bz2");
sh(`curl -fsSL ${JSON.stringify(`https://github.com/pyodide/pyodide/releases/download/${PYODIDE_VERSION}/pyodide-${PYODIDE_VERSION}.tar.bz2`)} -o ${JSON.stringify(pyTar)}`);
sh(`tar xjf ${JSON.stringify(pyTar)} -C ${JSON.stringify(tmp)}`);
mkdirSync(resolve(PUBLIC, "pyodide"), { recursive: true });
sh(`cp -r ${JSON.stringify(resolve(tmp, "pyodide") + "/.")} ${JSON.stringify(resolve(PUBLIC, "pyodide"))}`);

// Tidy: drop any stray sdists.
for (const f of readdirSync(PUBLIC)) {
  if (f.endsWith(".tar.gz")) rmSync(resolve(PUBLIC, f));
}
rmSync(tmp, { recursive: true, force: true });

console.log("\n[vendor] Done. Vendored (same-origin):");
console.log(`  - public/vendor/{worker_bootstrap.js,bridge.js,coi-serviceworker.js}`);
console.log(`  - public/${out(`ls ${JSON.stringify(PUBLIC)} | grep pyspark_connect_web || true`)}`);
console.log(`  - public/${want}`);
console.log(`  - public/pyodide/ (Pyodide ${PYODIDE_VERSION})`);
console.log("\n  Next: npm run build  (public/ -> dist/), then docker compose -f deploy/compose.yaml up -d --wait");
