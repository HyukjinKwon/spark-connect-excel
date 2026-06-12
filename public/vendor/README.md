<!-- SPDX-License-Identifier: Apache-2.0 -->

# public/vendor - reused pyspark-connect-web browser glue

These files are copied **verbatim** (Apache-2.0) from
[`pyspark-connect-web`](https://github.com/HyukjinKwon/pyspark-client-wasm) and
must be served **same-origin** (they cannot come from a CDN). Do not edit them
here - fix upstream and re-copy.

| File | Upstream path | Role |
|------|---------------|------|
| `worker_bootstrap.js` | `pyspark_connect_web/worker/worker_bootstrap.js` | Web Worker entry: loads Pyodide, micropip-installs the wheel, allocates the SAB. |
| `bridge.js` | `pyspark_connect_web/worker/bridge.js` | Main-thread `fetch` + SAB writeback for the Atomics handshake. |
| `coi-serviceworker.js` | `pyspark_connect_web/jupyterlite/coi-serviceworker.js` | COI shim for header-less hosts (task-pane optimization path only). |

The dialog host (`src/dialog/`) drives `worker_bootstrap.js` directly. By
default the bootstrap loads Pyodide + the wheels **same-origin** (`/pyodide/`,
`/pyspark_client-4.1.2-py3-none-any.whl`, `/pyspark_connect_web-0.2.0-py3-none-any.whl`)
- a CDN does not work under COI. Those large assets are git-ignored and vendored
separately (see `docs/reuse.md`); `src/runtime/pyodideHost.ts` only injects
`self.PCW_PYODIDE_INDEX_URL` / `self.PCW_WHEEL_URL` when explicitly overridden.

Synced from pyspark-connect-web `main` (0.2.0). Re-sync with
`npm run vendor:runtime`; this is the exact glue the end-to-end CI gate uses.
