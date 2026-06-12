<!-- SPDX-License-Identifier: Apache-2.0 -->

# End-to-End Test Harness

This directory contains the Playwright end-to-end tests for spark-connect-excel.

## What runs automatically (no Excel or Spark required)

### `coi.spec.ts` - the COI gate (LIVE)

The most important test in the repo. Asserts:

1. `self.crossOriginIsolated === true` in the dialog page
2. `typeof SharedArrayBuffer !== 'undefined'`
3. `typeof Atomics !== 'undefined'`
4. The `Cross-Origin-Opener-Policy: same-origin` response header is present
5. The `Cross-Origin-Embedder-Policy: credentialless` response header is present

These checks can run against the built `dist/` directory using `vite preview`
(which sets the COI headers defined in `vite.config.ts`). No Excel host, no Spark
cluster, and no Office.js are required. The dialog page may log JavaScript errors
about Office.js not loading (expected - there is no Office host), but the COI
header test passes before any scripts run.

Why this test matters: pyspark-connect-web's `SharedArrayBuffer`-based blocking
bridge (the backbone of blocking `.collect()`) requires cross-origin isolation.
If this test fails, nothing downstream works. It is the Phase 0 go/no-go gate.

### Running the COI gate locally

```bash
# 1. Build the add-in (produces dist/ with the correct headers):
npm run build

# 2. Run the COI gate:
npx playwright test coi.spec.ts

# Or via the npm script (runs all e2e tests):
npm run test:e2e
```

You must run `npm run build` before the tests; the webServer in
`playwright.config.ts` runs `npm run preview` which serves `dist/`. If `dist/`
is missing or stale, the tests will fail with a "page not found" error.

### Playwright browser

Only Chromium is configured (see `playwright.config.ts`). This is intentional:

- Office add-ins on Windows use **WebView2** (Chromium-based).
- Excel on the web runs in **Chrome or Edge** (Chromium-based).
- `COEP: credentialless` is a Chromium feature; it is not supported in Firefox
  or Safari (those hosts are not targeted by this add-in).

---

## What is deferred (requires Excel + Spark)

### `query-flow.spec.ts` - the Spark/Excel matrix (DEFERRED)

These tests cover the full connect -> runSQL -> range -> chart -> refresh flow
against a real Spark Connect server and a real Excel host. They are:

- Marked `test.fixme` in the Playwright report.
- Guarded by the `E2E_REQUIRE_STACK=true` environment variable.
- Skipped in CI (`ci.yml` and `e2e.yml` do not set `E2E_REQUIRE_STACK`).

**Why deferred:** There is no standard Playwright mechanism for driving an Excel
add-in task pane. The task pane runs inside a WebView2 (Windows) or WKWebView
(Mac) managed by Excel, not a standalone browser process. A custom harness using
the Office Test Framework or a recorded macro-based approach would be needed.
The COI gate covers the critical infrastructure; the full flow tests are staged
for manual verification + future CI enhancement.

### Manual verification steps

To run the full flow test matrix manually:

```bash
# 1. Bring up the Spark Connect + Envoy stack:
docker compose -f deploy/compose.yaml up
# Wait ~60s for Spark Connect to become healthy.

# 2. Build the add-in and start the preview server:
npm run build
npm run preview   # serves dist/ on http://localhost:3000

# 3. Sideload the add-in into Excel:
#    See scripts/sideload.md for platform-specific instructions.
npx office-addin-debugging start manifest.xml

# 4. Work through the manual checklist in query-flow.spec.ts.
```

### Running the deferred tests (if you have the stack)

```bash
E2E_REQUIRE_STACK=true npx playwright test query-flow.spec.ts
```

Note: the test bodies are stubs (`throw new Error("Not implemented")`). You
must implement the assertions using your harness before they can pass.

---

## CI integration

| Workflow | What runs |
|----------|-----------|
| `ci.yml` | Vitest unit tests + Python pytest + build + manifest validate (no Playwright) |
| `e2e.yml` | Build + install Chromium + `playwright test coi.spec.ts` |

The Playwright HTML report is uploaded as a CI artifact from `e2e.yml`.

---

## Playwright report

After a local run:

```bash
npx playwright show-report tests/e2e/playwright-report
```

---

## Running offline / without a Playwright browser download

The COI gate (`coi.spec.ts`) is **verified passing** in Chromium. If Playwright's
bundled browser isn't downloaded (restricted network), drive a system-installed
Chrome instead:

```bash
npm run build
PW_CHANNEL=chrome npm run test:e2e:coi
```

`PW_CHANNEL` is honored by `playwright.config.ts`; unset (CI) uses the bundled
Chromium. The five assertions are: `crossOriginIsolated === true`,
`SharedArrayBuffer` present, `Atomics` present, COOP `same-origin`, COEP
`credentialless` - i.e. the SharedArrayBuffer prerequisite the whole runtime
rests on.

## Adding new live tests

If a module is pure (no Office.js, no network) and automatable with `vite preview`,
add a new `*.spec.ts` in this directory. Keep it focused on the infrastructure
assertion that can run without Excel. Office-dependent assertions go into
`query-flow.spec.ts` as `deferredTest()` stubs until a harness is available.
