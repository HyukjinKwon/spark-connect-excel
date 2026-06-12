<!-- SPDX-License-Identifier: Apache-2.0 -->

# spark-connect-excel — Team Coordination

**Goal:** An Excel add-in that runs a **Spark SQL query against the user's own
Spark Connect cluster**, lands the result in a worksheet range, refreshes it, and
charts it — **with no backend server**, by hosting
[`pyspark-connect-web`](https://github.com/HyukjinKwon/pyspark-client-wasm)
(real PySpark Connect client, in-browser via Pyodide) inside the add-in.

Framing: **"Power Query, but the engine is your Spark cluster."** Surface = SQL.
Compute = pushed to Spark. Excel = last-mile renderer.

License: Apache-2.0 header on every source file. We **reuse** pyspark-connect-web
(consume the PyPI wheel; copy only its browser JS glue, unedited). See `PLAN.md`,
`DECISIONS.md`, `API_CONTRACT.md`.

## The frozen seam
`src/seam.ts` defines `SparkBridge` (the async API) + the task-pane⟷dialog message
envelope. Build against it; if you need a change, edit `src/seam.ts` AND append a
note here first.

## Module ownership (do NOT edit another lane's files)
| Lane | Area | Files (owned) | Status |
|------|------|---------------|--------|
| INTEGRATOR | scaffold, seam, configs, CI wiring | `*.md`, `package.json`, `tsconfig.json`, `vite.config.ts`, `eslint.config.js`, `vitest.config.ts`, `src/seam.ts`, `public/vendor/*` | scaffold done |
| A | Add-in shell & shared runtime | `manifest.xml`, `src/taskpane/taskpane.html`, `assets/*` | pending |
| B | COI gate & Dialog host | `src/dialog/dialog.html`, `src/dialog/dialogHost.ts`, `src/dialog/coi.ts` | pending |
| C | Pyodide + pcw runtime | `src/runtime/pyodideHost.ts`, `src/runtime/runPython.ts` | pending |
| D | SparkBridge seam impl | `src/bridge/sparkBridgeHost.ts`, `src/bridge/sparkBridgeClient.ts`, `python/spark_excel_runtime.py`, `python/tests/*` | pending |
| E | Query UX | `src/taskpane/taskpane.ts`, `src/taskpane/ui.css`, `src/taskpane/queryPanel.ts` | pending |
| F | Range writer | `src/excel/rangeWriter.ts`, `src/excel/typeFormat.ts` | pending |
| G | Refresh & binding | `src/excel/binding.ts`, `src/excel/refresh.ts` | pending |
| H | Charting | `src/excel/chart.ts` | pending |
| I | Connection/auth + deploy | `src/connection/connectionStore.ts`, `src/connection/connectionForm.ts`, `deploy/*` | pending |
| J | Packaging, e2e, docs | `tests/e2e/*`, `tests/unit/*`, `docs/*`, `.github/workflows/*`, `README.md`, `scripts/*` | pending |

## Conventions
- TypeScript strict; `npm run typecheck && npm run lint && npm test` must stay green.
- Apache-2.0 header on every source file.
- Tokens never written to a cell or document settings value (DECISIONS #6).
- Task pane must not import from `src/runtime/**` (DECISIONS #7).
- Need a new npm dep? Ask INTEGRATOR (owns `package.json`); don't edit it in a lane.
- Record non-obvious findings in `docs/findings-lane-<X>.md`.

## Notes log (newest last)
- INTEGRATOR 2026-06-12: Scaffolded repo; froze the seam in `src/seam.ts`
  (`SparkBridge` + message envelope). Locked v0 = SQL → range → refresh → chart,
  zero backend. COI host = Office Dialog window with COEP `credentialless`
  (DECISIONS #1/#2). Reuse pcw JS glue in `public/vendor/` (unedited). Lanes:
  claim your row and build against the seam. Hardest seam: D's `SparkBridgeHost`
  ⟷ C's `__pcwRunPython` ⟷ the Python runtime — agree on JSON shapes here before
  diverging (they're pinned in `API_CONTRACT.md` §3).
- INTEGRATOR 2026-06-12: All 10 lanes landed and integrated. Fixes applied at the
  seam: Office dialog event-handler union typing (DialogMessageReceived /
  DialogEventReceived); `OfficeRuntime.storage` reached via `globalThis` (no
  ambient-global dependency); `/vendor/bridge.js` imported via a non-literal
  specifier; manifest element order (Requirements before DefaultSettings) + Group
  `<Icon>` + `<?xml?>` first; icons moved to `public/assets/`. Added
  `tests/unit/sparkBridgeHost.test.ts` (fake RuntimeHost). **Green:** tsc,
  eslint+prettier, 208 vitest, 33 pytest, `vite build` (both entry points),
  `office-addin-manifest validate` (all Excel platforms), COI headers verified on
  the built dialog page (COOP same-origin + COEP credentialless). **Deferred (needs
  real Excel + a Spark Connect cluster, documented in `tests/e2e/README.md`):** the
  in-Excel Office/Spark e2e matrix and a headless-browser `crossOriginIsolated`
  assertion (Chromium download was unavailable locally; CI `e2e.yml` runs it).
- INTEGRATOR 2026-06-12 (finishing pass): Closed 2 of the 3 deferred items.
  (1) **COI gate now verified in a real browser** — ran `coi.spec.ts` against
  system Chrome (`PW_CHANNEL=chrome`, added to `playwright.config.ts`): all 5
  assertions pass (`crossOriginIsolated===true`, SharedArrayBuffer, Atomics, COOP
  same-origin, COEP credentialless). The architecture's make-or-break prerequisite
  is empirically proven, not just header-inferred. (2) **Production manifest
  generator** `scripts/build-manifest.mjs` + `npm run build:manifest -- --origin
  https://…` substitutes the dev origin and validates (closes the localhost
  placeholder). Fixed `test:e2e` to point at the config; added `test:e2e:coi`.
  **Still genuinely deferred (cannot run here):** the in-Excel connect→range→chart
  matrix — it needs a real Excel host + a running Spark Connect cluster + the
  pyspark-connect-web wheel published to PyPI (micropip installs it by name). Once
  the wheel is on PyPI and a cluster is up, sideload + run `query-flow.spec.ts`
  shape manually per `tests/e2e/README.md`.
