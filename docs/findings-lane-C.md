<!-- SPDX-License-Identifier: Apache-2.0 -->

# findings-lane-C.md - Pyodide + pcw Runtime Wiring

Lane C: `src/runtime/pyodideHost.ts` + `src/runtime/runPython.ts`.

---

## 1. Worker creation - Blob shim (classic worker)

`worker_bootstrap.js` calls `importScripts(...)` on line 63, which is only
available in **classic** (non-module) workers. Creating it with
`{type:"module"}` would fail at that call.

More critically, `worker_bootstrap.js` reads `self.PCW_PYODIDE_INDEX_URL` and
`self.PCW_WHEEL_URL` as **top-level variable initialisers** (lines 24-38):

```js
const PYODIDE_INDEX_URL =
  self.PCW_PYODIDE_INDEX_URL || "https://cdn.jsdelivr.net/...";
const MICROPIP_PKGS = [
  ...
  self.PCW_WHEEL_URL || "pyspark_connect_web-0.0.1.dev0-py3-none-any.whl",
];
```

These run the moment the script is parsed. We cannot inject globals after the
script has started. The clean solution is a **Blob shim**:

```js
// Generated at runtime by buildBootstrapBlob():
"use strict";
self.PCW_PYODIDE_INDEX_URL = "https://cdn.jsdelivr.net/pyodide/v0.28.0/full/"; // optional
self.PCW_WHEEL_URL = "pyspark-connect-web"; // micropip spec or explicit URL
importScripts('/vendor/worker_bootstrap.js');
```

This Blob is turned into an object URL and passed to `new Worker(blobUrl)` - no
`{type:"module"}` - giving a classic worker whose first (and only) script sets
the globals, then delegates to the real bootstrap.

If `pyodideIndexUrl` is not overridden in `BootOptions`, the shim omits that
line entirely, and `worker_bootstrap.js` falls through to its own default
(`https://cdn.jsdelivr.net/pyodide/v0.28.0/full/`).

Default `wheelUrl` is `"pyspark-connect-web"`, which micropip resolves from
PyPI (DECISIONS #3). Callers may pass a local wheel URL (e.g.
`"/vendor/pyspark_connect_web-0.0.1.dev0-py3-none-any.whl"`) via `BootOptions`.

---

## 2. bridge.js wiring (main thread)

`public/vendor/bridge.js` is an ES module that exports a single function:

```js
export function installBridge(worker) { ... }
```

`installBridge(worker)` creates a `Bridge` instance and attaches a `"message"`
listener to the worker. It handles two message types automatically:

- `{type:"pcw_sab", control, data}` - calls `bridge.attach(control, data)`,
  creating `Int32Array` / `Uint8Array` views over the SharedArrayBuffers that
  the worker allocated at boot. The worker re-posts this when it reallocates the
  data SAB for a larger payload.
- `{type:"pcw_rpc"}` - nudge from the worker that a request is written into the
  SAB. `bridge.handleRpc()` reads the header+body from the SAB, performs the
  real `fetch()` on the main thread, and writes the response back into the SAB,
  then calls `Atomics.notify` to wake the blocked Python thread.

Since `bridge.js` is served at `/vendor/bridge.js` (Vite serves `public/` at
`/`), we load it with a dynamic `import()`:

```ts
const bridgeMod = await import("/vendor/bridge.js");
bridgeMod.installBridge(worker);
```

The `/* @vite-ignore */` comment suppresses Vite's static-analysis warning for
the non-static import path. Vite does not bundle files from `public/`; the
browser fetches `/vendor/bridge.js` directly at runtime.

This must happen **before** `pcw_boot` is posted, so bridge.js's listener is
attached before the worker sends `pcw_sab`. Order in `boot()`:

1. Create Worker from Blob shim.
2. `await import("/vendor/bridge.js")` -> `installBridge(worker)`.
3. Create `RpcHandle` (wires `pcw_result` / `pcw_run_error` listener).
4. Post `{type:"pcw_boot"}`.
5. Await `pcw_ready` / `pcw_error`.

---

## 3. PCW_WHEEL_URL / PCW_PYODIDE_INDEX_URL injection

| Variable | Injected as | Default |
|---|---|---|
| `PCW_PYODIDE_INDEX_URL` | `self.PCW_PYODIDE_INDEX_URL = "...";` in Blob shim | Worker's own default (`cdn.jsdelivr.net/pyodide/v0.28.0/full/`) when not in `BootOptions` |
| `PCW_WHEEL_URL` | `self.PCW_WHEEL_URL = "...";` in Blob shim | `"pyspark-connect-web"` (micropip PyPI resolution) |

Both are injected into the Blob source as JSON-stringified literals, so they
are set before `importScripts('/vendor/worker_bootstrap.js')` runs and before
the top-level initialisers of `worker_bootstrap.js` execute.

---

## 4. pcw_run / pcw_result / pcw_run_error protocol

`src/runtime/runPython.ts` implements the `RpcHandle` class, mirroring
`installRunPython(worker)` from
`pyspark_connect_web/jupyterlite/run_python_bridge.js`:

- `RpcHandle.send(code)` increments a sequence counter, stores a
  `{resolve, reject}` pair in a `Map<number, PendingCall>`, and posts
  `{type:"pcw_run", id, code}` to the worker.
- The worker responds with `{type:"pcw_result", id, result}` or
  `{type:"pcw_run_error", id, message}`; the listener resolves/rejects the
  matching promise.
- `RpcHandle.rejectAll(reason)` drains in-flight calls on worker termination.
- `RpcHandle.detach()` removes the message listener to avoid leaks.

`PyodideHost.runPython(src)` delegates directly to `this._rpc.send(src)`.

---

## 5. Idempotency and error handling

- `boot()` is idempotent: a second call after success returns immediately.
- A concurrent second `boot()` call (while booting) throws immediately to
  prevent races (callers should use a single shared instance).
- On any boot failure (worker error event OR `pcw_error` message OR promise
  rejection in the dynamic import), `_cleanup()` terminates the worker, rejects
  all pending RPC calls, and revokes the Blob URL.
- `terminate()` calls `_cleanup()`, ensuring all pending `runPython` callers
  get a rejection rather than hanging.

---

## 6. COI precondition

`worker_bootstrap.js` calls `assertIsolated()` inside `boot()` (before
allocating SABs), which throws if `SharedArrayBuffer` is undefined or
`self.crossOriginIsolated !== true`. Lane B guarantees the dialog page is
served with `COOP: same-origin` + `COEP: credentialless` (DECISIONS #1, #2),
so the worker inherits isolation. PyodideHost does not re-check; the worker's
`pcw_error` message surfaces the failure naturally.
