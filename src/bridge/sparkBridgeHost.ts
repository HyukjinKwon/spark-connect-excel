// SPDX-License-Identifier: Apache-2.0
//
// sparkBridgeHost.ts — dialog-side implementation of SparkBridge.
//
// Runs inside the COI Office Dialog window.  Delegates real work to the Python
// runtime via RuntimeHost.runPython().  The Python module is bundled at build
// time via the `?raw` import and exec()'d into the Pyodide global namespace.

import type {
  SparkBridge,
  ConnectOptions,
  SparkResult,
  ColumnMeta,
  RuntimeHost,
  RuntimeStatus,
} from "../seam";
import { parseConnectResult, parseResult, parseSchema } from "./marshal";

// Vite/esbuild ?raw import — the Python source is inlined as a string.
import runtimeSource from "../../python/spark_excel_runtime.py?raw";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Build a Python snippet that safely passes `args` (as a JSON literal) into a
 * module-level function call:
 *
 *   __pcw_args = json.loads(r'''<json>''')
 *   <fn>(*__pcw_args)
 *
 * Using a raw triple-quoted string avoids any Python string-escaping issues
 * with the SQL text.  The result of the last expression is what runPython
 * returns (the JSON string produced by the function).
 */
function callPy(fn: string, ...args: unknown[]): string {
  // Encode args as JSON.  Triple-single-quote is safe here because JSON never
  // produces a lone single-quote followed by two more; but to be extra safe we
  // pick triple-single-quote and ensure the JSON doesn't contain the delimiter.
  // JSON never contains `'''` so this is always safe.
  const argsJson = JSON.stringify(args);
  return [
    `import json as __json`,
    `__pcw_args = __json.loads(r'''${argsJson}''')`,
    `${fn}(*__pcw_args)`,
  ].join("\n");
}

// ---------------------------------------------------------------------------
// SparkBridgeHost
// ---------------------------------------------------------------------------

export class SparkBridgeHost implements SparkBridge {
  private readonly _host: RuntimeHost;
  private _ready = false;
  private _connected = false;
  private _bootPromise: Promise<void> | null = null;

  constructor(host: RuntimeHost) {
    this._host = host;
  }

  // -------------------------------------------------------------------------
  // ensureReady
  // -------------------------------------------------------------------------

  /**
   * Boot Pyodide (via RuntimeHost), install pyspark-connect-web, then load the
   * Python runtime module by exec()-ing its source into the Pyodide globals.
   *
   * Idempotent: subsequent calls return the same promise / resolve immediately.
   */
  ensureReady(): Promise<void> {
    if (this._bootPromise !== null) {
      return this._bootPromise;
    }
    this._bootPromise = this._doEnsureReady();
    return this._bootPromise;
  }

  private async _doEnsureReady(): Promise<void> {
    // 1. Boot Pyodide + install the pcw wheel (Lane C / RuntimeHost).
    await this._host.boot();

    // 2. Install pyspark-connect-web (idempotent).
    await this._host.runPython(`import pyspark_connect_web as __pcw; __pcw.install(); "ok"`);

    // 3. Load the runtime module by exec()-ing its source into the global
    //    namespace so top-level functions (connect, run_sql, schema_of) are
    //    directly callable from subsequent runPython snippets.
    //
    //    We write the source to the Pyodide in-memory filesystem under
    //    /spark_excel_runtime.py and import it, which:
    //      - gives the module its own __name__ so relative imports work, and
    //      - makes the import idempotent (Python caches in sys.modules).
    //
    //    The source is embedded as a base64 string to avoid any quoting issues.
    const b64 = btoa(unescape(encodeURIComponent(runtimeSource)));
    const loaderSnippet = [
      `import base64 as __b64, sys as __sys`,
      `__src = __b64.b64decode(${JSON.stringify(b64)}).decode()`,
      `if "spark_excel_runtime" not in __sys.modules:`,
      `    import types as __types`,
      `    __mod = __types.ModuleType("spark_excel_runtime")`,
      `    exec(compile(__src, "spark_excel_runtime.py", "exec"), __mod.__dict__)`,
      `    __sys.modules["spark_excel_runtime"] = __mod`,
      `from spark_excel_runtime import connect, run_sql, schema_of`,
      `"loaded"`,
    ].join("\n");

    await this._host.runPython(loaderSnippet);
    this._ready = true;
  }

  // -------------------------------------------------------------------------
  // connect
  // -------------------------------------------------------------------------

  async connect(uri: string, opts?: ConnectOptions): Promise<void> {
    await this.ensureReady();
    const token: string | null = opts?.token ?? null;
    const snippet = callPy("connect", uri, token);
    const raw = await this._host.runPython(snippet);
    parseConnectResult(raw);
    this._connected = true;
  }

  // -------------------------------------------------------------------------
  // runSQL
  // -------------------------------------------------------------------------

  async runSQL(sql: string, rowCap: number): Promise<SparkResult> {
    await this.ensureReady();
    const snippet = callPy("run_sql", sql, rowCap);
    const raw = await this._host.runPython(snippet);
    return parseResult(raw);
  }

  // -------------------------------------------------------------------------
  // schemaOf
  // -------------------------------------------------------------------------

  async schemaOf(sql: string): Promise<ColumnMeta[]> {
    await this.ensureReady();
    const snippet = callPy("schema_of", sql);
    const raw = await this._host.runPython(snippet);
    return parseSchema(raw);
  }

  // -------------------------------------------------------------------------
  // status  (synchronous, safe to poll)
  // -------------------------------------------------------------------------

  status(): RuntimeStatus {
    return {
      // In the dialog window we own the origin, so crossOriginIsolated is
      // available directly.  Fallback to false if not defined (SSR / tests).
      crossOriginIsolated: typeof self !== "undefined" && self.crossOriginIsolated === true,
      pyodideReady: this._ready,
      connected: this._connected,
    };
  }

  // -------------------------------------------------------------------------
  // cancel  (best-effort, non-blocking)
  // -------------------------------------------------------------------------

  cancel(): void {
    // pyspark-connect-web does not expose an interrupt API at the Python level.
    // The best we can do is attempt to stop the Spark session; this is a no-op
    // if no query is in flight and may not cancel a blocking toPandas() call.
    // A full cancel would require the SAB/Atomics interrupt path (Lane C).
    if (this._connected) {
      void this._host
        .runPython(
          // Stop the session gracefully; ignore any error.
          [
            `try:`,
            `    from spark_excel_runtime import _spark as __s`,
            `    if __s is not None: __s.stop()`,
            `except Exception:`,
            `    pass`,
            `"cancelled"`,
          ].join("\n"),
        )
        .catch(() => {
          /* best-effort — ignore */
        });
      this._connected = false;
    }
  }
}
