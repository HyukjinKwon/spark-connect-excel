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
   the cross-origin office.js script without that origin sending CORP headers.
   Chromium-based Excel hosts (WebView2 / Edge / Chrome) support it. The dev
   server and the deploy static host both send `COOP: same-origin` +
   `COEP: credentialless`. CI guards their presence. NOTE: credentialless does
   NOT make a cross-origin CDN usable for the worker's `importScripts` of
   Pyodide - per pyspark-connect-web that is blocked in Chromium regardless. So
   Pyodide and the wheels are served SAME-ORIGIN (see #3).

3. **We reuse pyspark-connect-web; we do not fork it.** We consume the published
   package and copy only its browser JS glue (`worker_bootstrap.js`, `bridge.js`,
   `coi-serviceworker.js`) into `public/vendor/` verbatim (re-synced from upstream;
   provenance + pinned commit in `docs/reuse.md`). The heavy runtime is served
   **same-origin** next to the app (these come from pyspark-connect-web's build,
   version-matched; vendored into `public/`, git-ignored - see `docs/reuse.md`):
   - `/pyodide/` - the Pyodide distribution (a CDN does not work under COI),
   - `/pyspark-4.0.0-py2.py3-none-any.whl` - PySpark is sdist-only on PyPI,
   - `/pyspark_connect_web-*.whl` - the pcw wheel.
   `micropip` still fetches the small pure deps (`protobuf`, `googleapis-common-protos`,
   `py4j`) from PyPI at runtime. Override any URL via `self.PCW_PYODIDE_INDEX_URL` /
   `self.PCW_PYSPARK_WHEEL_URL` / `self.PCW_WHEEL_URL`.

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
