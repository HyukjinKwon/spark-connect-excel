<!-- SPDX-License-Identifier: Apache-2.0 -->

# Lane J findings — Packaging, Distribution, E2E, and Docs

**Date:** 2026-06-12

---

## What is tested live vs. deferred

### Live (automated, no Excel or Spark required)

| Test | File | What it asserts |
|------|------|----------------|
| COI gate — crossOriginIsolated | `tests/e2e/coi.spec.ts` | `self.crossOriginIsolated === true` against built `dist/` via `vite preview` |
| SAB availability | `tests/e2e/coi.spec.ts` | `typeof SharedArrayBuffer !== 'undefined'` |
| Atomics availability | `tests/e2e/coi.spec.ts` | `typeof Atomics !== 'undefined'` |
| COOP header | `tests/e2e/coi.spec.ts` | `Cross-Origin-Opener-Policy: same-origin` in response headers |
| COEP header | `tests/e2e/coi.spec.ts` | `Cross-Origin-Embedder-Policy: credentialless` in response headers |
| encodeMessage/decodeMessage | `tests/unit/seam.test.ts` | Round-trip for req/res/evt; SC_URI_HINT shape |
| normalizeRemoteUri | `tests/unit/seam.test.ts` | http/https scheme, host:port format |
| parseResult/parseSchema | `tests/unit/marshal.test.ts` | Valid JSON to typed objects; {ok:false} throws |
| parseConnectResult | `tests/unit/marshal.test.ts` | {ok:true} is void; {ok:false} throws |
| buildRemoteUri | `tests/unit/connectionStore.test.ts` | sc://host:port/;transport=grpcweb form |
| saveConnection/loadConnection | `tests/unit/connectionStore.test.ts` | Round-trip via memorySettingsBackend |
| Token helpers | `tests/unit/connectionStore.test.ts` | saveToken/loadToken/clearToken never write to SettingsBackend |
| SparkBridgeClient send | `tests/unit/sparkBridgeClient.test.ts` | runSQL/connect/ensureReady emit correct BridgeRequest |
| SparkBridgeClient resolve | `tests/unit/sparkBridgeClient.test.ts` | BridgeResponse resolves promise |
| SparkBridgeClient reject | `tests/unit/sparkBridgeClient.test.ts` | Error BridgeResponse rejects promise |
| status() | `tests/unit/sparkBridgeClient.test.ts` | Synchronous; updates from BridgeEvent "status" |
| cancel() | `tests/unit/sparkBridgeClient.test.ts` | Sends req, fire-and-forget |
| rejectAll() | `tests/unit/sparkBridgeClient.test.ts` | Rejects all pending calls |
| COI header guard | `scripts/check-coi-headers.sh` | vite.config.ts + deploy/ have COOP/COEP |

### Deferred (requires Excel + Spark stack)

| Test | File | Reason deferred |
|------|------|----------------|
| connect to runSQL to range | `tests/e2e/query-flow.spec.ts` | Requires sideloaded Excel add-in + Spark cluster |
| Truncation banner | `tests/e2e/query-flow.spec.ts` | Same |
| Refresh (range rewrite) | `tests/e2e/query-flow.spec.ts` | Same |
| Insert Chart | `tests/e2e/query-flow.spec.ts` | Same |
| Time-series line chart inference | `tests/e2e/query-flow.spec.ts` | Same |
| Token never in document settings | `tests/e2e/query-flow.spec.ts` | Requires Office.js + workbook inspection |
| Query binding survives reopen | `tests/e2e/query-flow.spec.ts` | Requires save/reopen cycle |
| cancel() interrupts in-flight query | `tests/e2e/query-flow.spec.ts` | Same |

---

## Cross-lane integration issues found

The following issues were identified during the inventory. Lane J does not edit other lanes' files. Reported to the integrator.

### Issue J-1: normalizeRemoteUri produces http(s):// but buildRemoteUri produces sc://

**Severity:** Documentation discrepancy (not a runtime bug)

`src/seam.ts#normalizeRemoteUri(host, port, tls)` returns `http(s)://host:port` (the HTTP shorthand accepted by pyspark-connect-web). Meanwhile, `src/connection/connectionStore.ts#buildRemoteUri(cfg)` returns `sc://host:port/;transport=grpcweb`.

The comment in `connectionStore.ts` says "This is consistent with `seam.ts#normalizeRemoteUri`" but they produce different URI schemes. The `SC_URI_HINT` constant in `seam.ts` is the sc:// form; the hint and `normalizeRemoteUri` are inconsistent in scheme.

**Recommendation:** Clarify in `seam.ts` that `normalizeRemoteUri` is the HTTP shorthand alternative, or rename it to avoid confusion with the primary sc:// path used by `buildRemoteUri` and the Python runtime.

### Issue J-2: SparkBridgeClient.runSQL and schemaOf have redundant async keyword

**Severity:** Minor (TypeScript style, no runtime impact)

Both `runSQL` and `schemaOf` are declared `async` but directly return `this._call(...)` without any `await` inside. The `async` keyword is redundant. Not a bug but worth cleaning up.

### Issue J-3: SparkBridgeHost.cancel() sets connected=false unconditionally

**Severity:** Minor behaviour gap

`SparkBridgeHost.cancel()` sets `this._connected = false` even when the best-effort cancel attempt fails or is a no-op (e.g. no query in flight). The `SparkBridgeClient` does not push a status event after cancel, so the task pane status display may be stale.

**Recommendation:** Lane D to decide: either push a status event after cancel, or document that `status().connected` is unreliable after `cancel()`.

### Issue J-4: spark_excel_runtime.py line 207 has dead-code conditional

**Severity:** Minor code smell (Python)

```python
sep = ";" if ";" in uri else ";"
```

Both branches produce `";"`. The conditional is dead code. No runtime impact. Reported to Lane D.

---

## Unit test coverage gaps (by design)

The following modules were not unit-tested by Lane J because they cannot be tested purely:

| Module | Reason |
|--------|--------|
| `src/excel/rangeWriter.ts` | Requires `Excel.run` / Office.js mock |
| `src/excel/refresh.ts` | Calls writeResult (Office.js) |
| `src/taskpane/queryPanel.ts` | DOM + Office.js |
| `src/dialog/dialogHost.ts` | Boot + Office.js message bus |
| `src/runtime/pyodideHost.ts` | Web Worker + Blob URL |
| `src/runtime/runPython.ts` | Wraps Worker postMessage |
| `src/dialog/coi.ts` | `self.crossOriginIsolated` covered by live e2e COI gate |

These gaps are intentional. The COI gate covers the most important infrastructure check. Pure modules are unit-tested. Office-dependent modules are deferred to the full e2e matrix.

---

## Summary of files created

| File | Purpose |
|------|---------|
| `tests/unit/seam.test.ts` | encodeMessage/decodeMessage; normalizeRemoteUri; SC_URI_HINT |
| `tests/unit/marshal.test.ts` | parseResult / parseSchema / parseConnectResult |
| `tests/unit/connectionStore.test.ts` | buildRemoteUri; save/load round-trip; token isolation |
| `tests/unit/sparkBridgeClient.test.ts` | Full SparkBridgeClient with fake transport |
| `tests/e2e/playwright.config.ts` | Playwright config (vite preview webServer, Chromium) |
| `tests/e2e/coi.spec.ts` | Live COI gate |
| `tests/e2e/query-flow.spec.ts` | Deferred Spark/Excel matrix (test.fixme) |
| `tests/e2e/README.md` | E2E harness documentation |
| `.github/workflows/ci.yml` | Main CI pipeline |
| `.github/workflows/e2e.yml` | E2E CI: COI gate + Playwright report artifact |
| `docs/architecture.md` | Full architecture, ASCII diagram, decision rationale |
| `docs/installation.md` | Prerequisites, build, dev server, sideload |
| `docs/usage.md` | User guide |
| `docs/security.md` | Threat model |
| `docs/reuse.md` | pyspark-connect-web provenance |
| `docs/distribution.md` | Sideload vs AppSource, hosting requirements |
| `docs/findings-lane-J.md` | This file |
| `scripts/sideload.md` | Sideload quick-start |
| `scripts/check-coi-headers.sh` | COI header guard script |
| `README.md` | Root project README |
