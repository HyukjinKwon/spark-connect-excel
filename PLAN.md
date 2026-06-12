<!-- SPDX-License-Identifier: Apache-2.0 -->

# PLAN — Spark Connect for Excel

**One line:** An Excel add-in that runs a **Spark SQL query against your own Spark
Connect cluster, lands the result in a worksheet range, refreshes it, and charts
it — with _no backend server_, powered by [`pyspark-connect-web`](https://github.com/HyukjinKwon/pyspark-client-wasm).**

Think **"Power Query, but the engine is your Spark cluster."** Excel is the
last-mile UI; Spark does the compute; the query runs in-browser via Pyodide.

---

## Why this shape (and not a "PySpark IDE in Excel")

Excel users do not author DataFrame transformations — they pull data in, refresh
it, and chart it. So we surface **SQL**, not the Python API, even though
`pyspark-connect-web` exposes the full client underneath. The value proposition
is "your lakehouse table, in a cell, refreshable," not "write Spark code here."

The compute lives on the cluster: **push aggregation down, bring back the small
result.** A `groupBy().agg()` returns 12 rows, not a billion. Excel renders the
12.

---

## v0 scope (locked — keep it small)

v0 is **done** when, against a real Spark Connect server (Spark 4.x) behind the
Envoy grpc-web proxy, with **no backend service of our own**:

1. The add-in loads in Excel and reports `crossOriginIsolated === true`.
2. User enters a connection (`sc://host:port/;transport=grpcweb` + optional
   bearer token) and a **SQL query**.
3. Result lands in a worksheet **range** (header row + typed cells), with a
   row cap (default 10k) and a clear "truncated" indicator.
4. A **Refresh** button re-runs the bound query and rewrites the range.
5. **Insert chart** produces a native Excel chart from the result range.

**Out of v0 (explicitly):** writing PySpark/DataFrame code, the `=SPARK.SQL()`
custom function (Phase 4 stretch), schema-browser tree, parameterized
cell-driven queries, write-back to Spark. Don't build these in v0.

---

## Architecture

```
  Excel host (Windows WebView2 / Mac WKWebView / Excel-on-web iframe)
        │  Office.js shared runtime
        ▼
  Add-in task pane (UI: connection, SQL box, run/refresh/chart)
        │  SparkBridge (frozen seam, §"The seam")
        ▼
  COI host  ── EITHER the task pane itself (if crossOriginIsolated)
            └─ OR an Office Dialog window we fully control (fallback, §Risks)
        │  hosts Pyodide + pyspark-connect-web + SAB worker bridge  (REUSED from pcw)
        ▼
  pcw: spark.sql(userQuery).toPandas()      ← real PySpark Connect client, unchanged
        │  grpc-web over fetch
        ▼
  Envoy grpc_web proxy  (REUSED from pcw deploy/, CORS adapted for add-in origin)
        │  gRPC/HTTP2
        ▼
  user's Spark Connect server (Spark 4.x)
        ▲
  pandas → marshalled rows → Office.js → worksheet range + chart
```

**Reuse, do not fork.** `pyspark-connect-web` already provides the hard parts:
the in-browser PySpark client, the SAB blocking bridge (`worker/bridge.js`,
`worker/worker_bootstrap.js`), `__pcwRunPython`, the `coi-serviceworker.js` COI
shim, and the Envoy `deploy/` stack. This project is an **Excel host** around
that package — we add Office.js glue, not Spark plumbing.

---

## The seam (frozen — like pcw's API_CONTRACT.md)

The single stable interface between the **Excel-facing lanes** and the
**runtime lanes**. Excel lanes build against it; runtime lanes implement it. Do
not change a signature without a note in `COORDINATION.md`.

```ts
interface ColumnMeta { name: string; type: string; }            // type = Spark SQL type name
interface SparkResult {
  schema: ColumnMeta[];
  rows: unknown[][];          // already JS-native (numbers, strings, bools, ISO dates)
  rowCount: number;           // rows returned (post-cap)
  truncated: boolean;         // true if the cap clipped the result
}

interface SparkBridge {
  ensureReady(): Promise<void>;                       // Pyodide loaded + pcw.install() + COI verified
  connect(uri: string, opts?: { token?: string }): Promise<void>;
  runSQL(sql: string, rowCap: number): Promise<SparkResult>;
  schemaOf(sql: string): Promise<ColumnMeta[]>;       // EXPLAIN-only, no data (for chart-type inference)
  status(): { crossOriginIsolated: boolean; pyodideReady: boolean; connected: boolean };
  cancel(): void;
}
```

`runSQL` internally is just: `spark.sql(sql).limit(rowCap+1).toPandas()` →
detect overflow → marshal to `SparkResult`. The `+1` detects truncation.

If the COI host is an **Office Dialog window** (fallback), `SparkBridge` is the
same interface implemented over `postMessage` to that window.

---

## The 10 agents (lanes)

Mirrors pcw's lane model. A separate **Integrator/Coordinator** (not one of the
10) owns this PLAN, the seam, and `COORDINATION.md`, and sequences the gate.

| Lane | Owns | Deliverable |
|------|------|-------------|
| **A — Add-in shell** | `manifest.xml`, ribbon button, task-pane HTML/host, Office.js **shared runtime** config, CSP | Add-in installs (sideload) and opens a task pane in Excel |
| **B — COI gate & host strategy** ⚠️GATING | cross-origin-isolation inside Office | `crossOriginIsolated === true` in the runtime on web + Windows + Mac. Reuse `coi-serviceworker.js`; if the pane can't be isolated, deliver the **Office Dialog-window host**. **Go/No-Go owner.** |
| **C — Pyodide + pcw runtime** | Pyodide load, `micropip install pyspark-connect-web`, SAB worker wiring | `pcw.install()` + a live `SparkSession` inside the COI host; reuses pcw worker glue verbatim |
| **D — SparkBridge seam** | the frozen interface above + pandas→JS marshalling + (if needed) the Dialog `postMessage` layer | `runSQL()` returns a `SparkResult` end-to-end |
| **E — Query UX** | task-pane UI: connection fields, SQL editor, row-cap, destination, run/cancel, error display | A usable panel that drives `SparkBridge` |
| **F — Range writer** | pandas/`SparkResult` → worksheet range via Office.js: headers, type/number-format mapping, spill, truncation banner | Result appears correctly typed in cells |
| **G — Refresh & binding** | bind a range to its query (named range + stored query metadata in document settings), Refresh / Refresh-All | Re-running updates the same range in place |
| **H — Charting** | one-click native `charts.add` from the result range + chart-type inference from `schemaOf` | A chart appears bound to the result |
| **I — Connection/auth + deploy** | connection + **secure token handling**; adapt pcw `deploy/` Envoy CORS/CORP for the add-in origin; TLS+bearer prod overlay; one-command bring-up | `docker compose up` gives a cluster the add-in can reach |
| **J — Packaging, distribution, e2e & docs** | HTTPS-host the bundle (GitHub Pages), sideload + **AppSource** path, manifest validation, Playwright e2e (COI gate / query→range parity / chart render), user docs | Anyone can install and use it |

---

## Sequencing (phases)

**Phase 0 — GO / NO-GO (Lane B + thin Lane A).** Build the smallest add-in that
loads a hosted page and prints `crossOriginIsolated`. Try the task-pane path
with `coi-serviceworker.js`; if false on any host, stand up the Office
Dialog-window host and re-test. **Nothing else is worth building until one of
these is green on Windows, Mac, and web.** This is the whole project's risk.

**Phase 1 — Vertical slice (A, B, C, D, I).** "Type a SQL query in the task
pane, get rows back from a real Spark Connect server, with no backend." Proves
the seam end-to-end.

**Phase 2 — Excel features (E, F, G, H).** Build on the frozen seam in parallel:
range writing, refresh/binding, charts, polished query UX.

**Phase 3 — Ship (J).** Host the bundle, sideload + AppSource, e2e parity gate,
docs and hardening.

**Phase 4 — Stretch (post-v0).** `=SPARK.SQL("…")` streaming custom function,
schema browser, cell-parameterized queries.

---

## Top risks & mitigations

1. **⚠️ Can an Office.js runtime be cross-origin-isolated?** This is the make-or-
   break. SAB (and thus pcw's blocking `.collect()`) needs COOP/COEP. The pane
   may be a nested iframe whose embedder we don't control.
   - *Mitigation A:* `coi-serviceworker.js` (already in pcw) injects COI via a
     service worker + one reload — works on header-less hosts.
   - *Mitigation B (primary fallback):* host Pyodide+pcw in an **Office Dialog
     window** (`displayDialogAsync`) — a real top-level window we serve with our
     own COOP/COEP, so it _is_ isolatable. Task pane ↔ dialog via `postMessage`;
     the dialog owns Spark, the pane owns Excel. Same `SparkBridge`.
   - Phase 0 decides which path; everything downstream is identical.
2. **`office.js` CDN under `COEP: require-corp`.** Cross-origin scripts must send
   CORP or be loaded `crossorigin`. Mitigation: COEP `credentialless` mode, or
   self-host the office.js shim. Lane B owns proving this.
3. **Pyodide cold-start (seconds) + bundle size.** Acceptable for a *connector*
   (you query, then work); show a one-time "starting engine" state. Lane C/E.
4. **Large results.** Cap rows Excel-side (default 10k) and lean on pcw's
   bounded-window transfer; never `collect()` unbounded into a grid. Lane F.
5. **Token in a spreadsheet.** Keep bearer tokens in Office roaming settings /
   session only, never written to cells; enforcement is at the Envoy proxy.
   Lane I.

---

## What we are NOT building

- No backend compute service (that's the whole point — pcw runs client-side).
- No fork of PySpark or `pyspark-connect-web` (reuse the package as published to
  PyPI; contribute upstream if the Excel host surfaces gaps).
- No DataFrame/Python authoring UI in v0.

---

## Open decision for the maintainer

Confirm the **primary host strategy preference** if both work in Phase 0:
task-pane-isolated (simpler UX, one surface) vs. Dialog-window host (more
robust isolation, extra window). Default recommendation: **whichever Phase 0
proves first; prefer task-pane if it passes on all three hosts.**
