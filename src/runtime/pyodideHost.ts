// SPDX-License-Identifier: Apache-2.0
//
// pyodideHost.ts — implements RuntimeHost (src/seam.ts) by booting Pyodide +
// pyspark-connect-web inside a Web Worker (worker_bootstrap.js) and wiring the
// main-thread bridge (bridge.js) for the Atomics/SAB fetch-proxy handshake.
//
// Lane C owns this file. Do not import from src/taskpane/** (DECISIONS #7).
//
// ---- How the pieces fit together -------------------------------------------
//
//  1. Worker creation (Blob shim, classic worker)
//     worker_bootstrap.js is a CLASSIC worker (it calls importScripts, no
//     "type":"module"). It reads `self.PCW_PYODIDE_INDEX_URL` and
//     `self.PCW_WHEEL_URL` at parse time (before any async work), so we must
//     inject those globals before the script body runs.
//
//     We accomplish this by building a tiny Blob whose text is:
//       self.PCW_PYODIDE_INDEX_URL = "...";
//       self.PCW_WHEEL_URL = "...";
//       importScripts('/vendor/worker_bootstrap.js');
//     and creating `new Worker(blobUrl)` (no {type:"module"} → classic).
//
//  2. bridge.js wiring (main thread)
//     bridge.js is an ES module that exports `installBridge(worker)`. It
//     attaches a "message" listener to the worker; when the worker posts
//     {type:"pcw_sab"} bridge calls bridge.attach(control, data), and when
//     the worker posts {type:"pcw_rpc"} bridge calls handleRpc() which reads
//     the request from the SAB, performs a real fetch on the main thread, and
//     writes the response back — waking the worker's Atomics.wait.
//
//     Since bridge.js lives in public/vendor/ (served at /vendor/bridge.js),
//     we load it with a dynamic import() using its runtime URL. The import()
//     returns the module's exports; we call installBridge(worker) immediately.
//
//  3. Boot sequence
//     a. Build Blob shim, create classic Worker.
//     b. Dynamic import('/vendor/bridge.js') → installBridge(worker).
//     c. Post {type:"pcw_boot"} — worker begins loading Pyodide + micropip.
//     d. Listen for {type:"pcw_ready"} → resolve boot(); {type:"pcw_error"} →
//        reject. Progress is inferred from log messages where available.
//
//  4. runPython
//     Uses the pcw_run / pcw_result / pcw_run_error protocol via RpcHandle
//     (src/runtime/runPython.ts), mirroring installRunPython() in the upstream
//     run_python_bridge.js.

"use strict";

import type { BootOptions, RuntimeHost } from "../seam.js";
import { createRpc, type RpcHandle } from "./runPython.js";

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Build the classic-worker bootstrap Blob that injects URL globals then loads
 * the real worker_bootstrap.js via importScripts.
 *
 * The Blob approach is necessary because worker_bootstrap.js reads
 * `self.PCW_PYODIDE_INDEX_URL` / `self.PCW_WHEEL_URL` at script-parse time
 * (top-level variable initialisers), so we must set them before the script
 * runs. A classic Worker created from a Blob URL has the Blob as its only
 * script; the real bootstrap is pulled in via importScripts.
 */
function buildBootstrapBlob(
  pyodideIndexUrl: string | undefined,
  wheelUrl: string | undefined,
): string {
  // Only inject a global when explicitly overridden; otherwise worker_bootstrap.js
  // uses its same-origin defaults (/pyodide/, /pyspark_connect_web-*.whl). Per
  // pyspark-connect-web, Pyodide and the wheels MUST be same-origin: a cross-origin
  // CDN is blocked by COEP for the worker's importScripts even under credentialless.
  const line = (name: string, val: string | undefined) =>
    val !== undefined
      ? `self.${name} = ${JSON.stringify(val)};`
      : `// ${name} not overridden - worker_bootstrap.js uses its same-origin default`;

  // importScripts path is relative to the origin (Vite serves public/ at /).
  return [
    `"use strict";`,
    line("PCW_PYODIDE_INDEX_URL", pyodideIndexUrl),
    line("PCW_WHEEL_URL", wheelUrl),
    `importScripts('/vendor/worker_bootstrap.js');`,
  ].join("\n");
}

// ---------------------------------------------------------------------------
// PyodideHost
// ---------------------------------------------------------------------------

export class PyodideHost implements RuntimeHost {
  private _defaults: BootOptions;
  private _worker: Worker | null = null;
  private _rpc: RpcHandle | null = null;
  private _ready = false;
  private _booting = false;
  private _blobUrl: string | null = null;

  /**
   * @param defaults  Optional baseline BootOptions applied when boot() is
   *                  called without options (or with partial options).
   */
  constructor(defaults: BootOptions = {}) {
    this._defaults = defaults;
  }

  // --------------------------------------------------------------------------
  // RuntimeHost.ready
  // --------------------------------------------------------------------------

  get ready(): boolean {
    return this._ready;
  }

  // --------------------------------------------------------------------------
  // RuntimeHost.boot
  // --------------------------------------------------------------------------

  /**
   * Boot Pyodide + install pyspark-connect-web in the Web Worker.
   *
   * Idempotent: a second call while already booted resolves immediately.
   * A concurrent second call while booting is NOT supported (it will reject
   * with a guard error) — callers should ensure a single boot() invocation.
   *
   * @param opts        Override pyodideIndexUrl or wheelUrl.
   * @param onProgress  Optional human-readable progress callback. Fired with
   *                    coarse messages ("Booting worker…", "Pyodide ready",
   *                    "pyspark-connect-web installed").
   */
  async boot(opts?: BootOptions, onProgress?: (msg: string) => void): Promise<void> {
    if (this._ready) return; // idempotent
    if (this._booting) {
      throw new Error("PyodideHost: boot() called while already booting");
    }
    this._booting = true;

    const merged: BootOptions = { ...this._defaults, ...opts };
    // Leave both undefined by default so worker_bootstrap.js uses its same-origin
    // defaults (Pyodide at /pyodide/, wheels at the site root). Only override when
    // the caller hosts them elsewhere (same-origin).
    const wheelUrl = merged.wheelUrl; // undefined -> worker default (same-origin wheel)
    const pyodideIndexUrl = merged.pyodideIndexUrl; // undefined -> worker default (/pyodide/)

    try {
      // ---- 1. Build the Blob shim and create a classic Worker ---------------
      onProgress?.("Booting Pyodide worker…");
      const shimSrc = buildBootstrapBlob(pyodideIndexUrl, wheelUrl);
      const blob = new Blob([shimSrc], { type: "application/javascript" });
      this._blobUrl = URL.createObjectURL(blob);

      // Classic worker (no {type:"module"}) — worker_bootstrap.js uses
      // importScripts, which is only available in classic workers.
      const worker = new Worker(this._blobUrl);
      this._worker = worker;

      // ---- 2. Wire bridge.js on the main thread ----------------------------
      // bridge.js is an ES module served at /vendor/bridge.js. Dynamic import
      // works because the dialog page itself is a module-capable context.
      // installBridge attaches a "message" listener to the worker that handles
      // pcw_sab (attach SAB views) and pcw_rpc (perform fetch + SAB writeback).
      onProgress?.("Wiring SAB bridge…");
      // Non-literal specifier: the bridge is a same-origin runtime asset under
      // public/vendor/, resolved by the browser at execution time, not bundled.
      const bridgeModUrl = "/vendor/bridge.js";
      const bridgeMod = (await import(/* @vite-ignore */ bridgeModUrl)) as {
        installBridge(worker: Worker): void;
      };
      bridgeMod.installBridge(worker);

      // ---- 3. Wire the RPC handle for pcw_run / pcw_result -----------------
      this._rpc = createRpc(worker);

      // ---- 4. Send pcw_boot and wait for pcw_ready / pcw_error -------------
      await new Promise<void>((resolve, reject) => {
        // Reject and clean up on any unhandled worker error.
        const onError = (ev: ErrorEvent) => {
          cleanup();
          this._fail(reject, new Error(`Worker error during boot: ${ev.message ?? String(ev)}`));
        };
        const onMessage = (ev: MessageEvent) => {
          const msg = (ev.data ?? {}) as Record<string, unknown>;
          if (msg.type === "pcw_ready") {
            cleanup();
            onProgress?.("pyspark-connect-web ready");
            resolve();
          } else if (msg.type === "pcw_error") {
            cleanup();
            this._fail(reject, new Error(`pcw boot error: ${String(msg.message ?? "unknown")}`));
          }
          // Coarse progress inferred from boot stages: Pyodide logs come via
          // console inside the worker, which we cannot intercept here. We
          // surface a single "Loading Pyodide + packages…" message right after
          // pcw_boot is posted so the UI is not silent during the long wait.
        };
        const cleanup = () => {
          worker.removeEventListener("message", onMessage);
          worker.removeEventListener("error", onError);
        };

        worker.addEventListener("message", onMessage);
        worker.addEventListener("error", onError);

        onProgress?.("Loading Pyodide + packages (this may take a minute)…");
        worker.postMessage({ type: "pcw_boot" });
      });

      this._ready = true;
    } catch (err) {
      // On any boot failure: clean up worker resources.
      this._cleanup();
      this._booting = false;
      throw err;
    }

    this._booting = false;
  }

  // --------------------------------------------------------------------------
  // RuntimeHost.runPython
  // --------------------------------------------------------------------------

  /**
   * Execute Python source in the worker and resolve with the stringified last
   * expression. Implements the pcw_run / pcw_result / pcw_run_error protocol.
   *
   * @throws if boot() has not completed, or the Python code raises.
   */
  async runPython(src: string): Promise<string> {
    if (!this._ready || this._rpc === null) {
      throw new Error("PyodideHost: not ready — call boot() first");
    }
    return this._rpc.send(src);
  }

  // --------------------------------------------------------------------------
  // RuntimeHost.terminate
  // --------------------------------------------------------------------------

  /**
   * Tear down the worker and reject any in-flight runPython calls.
   */
  terminate(): void {
    this._cleanup();
  }

  // --------------------------------------------------------------------------
  // Private helpers
  // --------------------------------------------------------------------------

  /** Reject `reject` with `err` and mark the host not-ready. */
  private _fail(reject: (reason: unknown) => void, err: Error): void {
    this._ready = false;
    reject(err);
  }

  /** Clean up the worker, RPC handle, and Blob URL. */
  private _cleanup(): void {
    if (this._rpc) {
      this._rpc.rejectAll(new Error("PyodideHost terminated"));
      this._rpc.detach();
      this._rpc = null;
    }
    if (this._worker) {
      this._worker.terminate();
      this._worker = null;
    }
    if (this._blobUrl) {
      URL.revokeObjectURL(this._blobUrl);
      this._blobUrl = null;
    }
    this._ready = false;
    this._booting = false;
  }
}
