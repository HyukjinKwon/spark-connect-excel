<!-- SPDX-License-Identifier: Apache-2.0 -->

# DECISIONS.md - invariants (keep these green)

Locked architectural decisions. Each has a guard (test or CI check) where
practical. Changing one requires a note in `COORDINATION.md`.

1. **The COI host is a separate window, not the task pane.** `SharedArrayBuffer`
   (the backbone of pyspark-connect-web's blocking `.collect()`) needs cross-
   origin isolation, and we cannot rely on Excel-on-web isolating the task-pane
   iframe. So Pyodide + pcw live in an **Office Dialog window** (`displayDialogAsync`)
   that we serve with our own COOP/COEP. The task pane talks to it over the Office
   dialog message channel (`messageParent`/`messageChild`). The task-pane COI path
   (via `coi-serviceworker.js`) is a runtime-detected optimization, never required.

2. **COEP is `credentialless`, not `require-corp`.** This lets the dialog load
   cross-origin Pyodide (jsDelivr) and the PyPI wheel without those origins
   sending CORP headers. Chromium-based Excel hosts (WebView2 / Edge / Chrome)
   support it. The dev server and the deploy static host both send
   `COOP: same-origin` + `COEP: credentialless`. CI guards their presence.

3. **We reuse pyspark-connect-web; we do not fork or vendor its Python.** The
   wheel is consumed from PyPI via `micropip`. Only its small browser JS glue
   (`worker_bootstrap.js`, `bridge.js`, `coi-serviceworker.js`) is copied into
   `public/vendor/` (it must be served same-origin). Those copies carry an
   upstream-provenance header and are not edited.

4. **SQL is the product surface *in the Excel add-in*.** Excel users query; they
   do not author DataFrames, so the task pane exposes `runSQL` only. Pushdown is
   the norm - bring back the small aggregated result. The standalone **web demo**
   (`src/demo/`) is developer-facing and intentionally exposes **both** SQL and a
   Python mode (`host.runPython`, `spark` pre-bound) - that surface showcases the
   full PySpark client. The Excel add-in stays SQL-first.

5. **Results are row-capped Excel-side (default 10k).** `runSQL` runs
   `.limit(cap + 1)`; a returned `cap + 1`th row sets `truncated = true` and is
   dropped. We never `collect()` unbounded into a grid.

6. **Bearer tokens never touch a cell.** They live in Office document/roaming
   settings (or session memory) and are enforced at the Envoy proxy. The range
   stores only the query text + endpoint host, never the token.

7. **The task pane never imports Pyodide/pcw.** All Python/Spark code runs in the
   dialog. The task pane is pure Office.js + UI + the bridge client. CI guards
   that `src/taskpane/**` does not import from `src/runtime/**`.

8. **Marshalling happens in Python, parsing in TS.** The Python runtime returns a
   single JSON string shaped exactly like `SparkResult`; TS only `JSON.parse`s it.
   No ad-hoc type coercion straddles the boundary.
