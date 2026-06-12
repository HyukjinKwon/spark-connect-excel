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
 * Convert a UTF-8 string to a base64-encoded string without using the
 * deprecated `unescape(encodeURIComponent(...))` idiom.
 *
 * Uses TextEncoder (available in all modern browsers and Node 16+) to get the
 * UTF-8 byte array, then maps each byte to its char and calls btoa().
 */
function utf8ToBase64(str: string): string {
  const bytes = new TextEncoder().encode(str);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

// Sentinel that brackets the runtime's JSON payload in the worker's captured
// stdout. The worker (worker_bootstrap.js) captures stdout, NOT the Python
// return value, so the snippet must PRINT its result. We wrap it in this
// unique marker so the payload survives any incidental stdout (Pyodide /
// pyspark-connect-web warnings printed before or after the result).
const PCW_MARK = "<<<__SCX_PCW__>>>";

/**
 * Build a Python snippet that safely passes `args` into a module-level
 * function call via base64-encoded JSON, then PRINTS the function's JSON
 * result bracketed by PCW_MARK so the host can recover it from stdout:
 *
 *   __pcw_args = __json.loads(__b64.b64decode("<b64>").decode())
 *   __pcw_out = <fn>(*__pcw_args)
 *   print(MARK + __pcw_out + MARK, end="")
 *
 * Base64-encoding the JSON payload completely eliminates any risk of the
 * argument content (e.g. SQL with triple-quotes or backslashes) breaking
 * Python string parsing.
 */
function callPy(fn: string, ...args: unknown[]): string {
  const argsJson = JSON.stringify(args);
  const argsB64 = utf8ToBase64(argsJson);
  const mark = JSON.stringify(PCW_MARK);
  return [
    `import json as __json, base64 as __b64`,
    `__pcw_args = __json.loads(__b64.b64decode(${JSON.stringify(argsB64)}).decode())`,
    `__pcw_out = ${fn}(*__pcw_args)`,
    // A trailing newline is REQUIRED: the worker captures stdout via Pyodide's
    // setStdout({batched}), which only flushes the buffer on a newline. Without
    // it the printed payload never reaches the host (raw stays empty). The
    // newline lands outside the closing marker, so extractPayload is unaffected.
    `print(${mark} + (__pcw_out if isinstance(__pcw_out, str) else __json.dumps(__pcw_out)) + ${mark})`,
  ].join("\n");
}

/**
 * Recover the JSON payload from the worker's captured stdout. When the snippet
 * printed the PCW_MARK-bracketed payload, return exactly what is between the
 * markers (discarding any incidental stdout outside them). When no marker is
 * present (e.g. unit tests feeding canned JSON directly), fall back to the
 * trimmed raw string so callers stay backward compatible.
 */
function extractPayload(raw: string): string {
  const first = raw.indexOf(PCW_MARK);
  if (first === -1) return raw.trim();
  const second = raw.indexOf(PCW_MARK, first + PCW_MARK.length);
  if (second === -1) return raw.trim();
  return raw.slice(first + PCW_MARK.length, second);
}

// ---------------------------------------------------------------------------
// SparkBridgeHost
// ---------------------------------------------------------------------------

export class SparkBridgeHost implements SparkBridge {
  private readonly _host: RuntimeHost;
  private _ready = false;
  private _connected = false;
  private _bootPromise: Promise<void> | null = null;
  /** Memoised last-successfully-connected args. Reset by cancel(). */
  private _lastConnectUri: string | null = null;
  private _lastConnectToken: string | null = null;

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
    this._bootPromise = this._doEnsureReady().catch((err: unknown) => {
      // Reset so that a subsequent call can retry the boot sequence.
      this._bootPromise = null;
      this._ready = false;
      throw err;
    });
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
    const b64 = utf8ToBase64(runtimeSource);
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

    // Idempotent: skip the Python connect() if already connected with identical args.
    if (this._connected && this._lastConnectUri === uri && this._lastConnectToken === token) {
      return;
    }

    const snippet = callPy("connect", uri, token);
    const raw = await this._host.runPython(snippet);
    parseConnectResult(extractPayload(raw));
    this._connected = true;
    this._lastConnectUri = uri;
    this._lastConnectToken = token;
  }

  // -------------------------------------------------------------------------
  // runSQL
  // -------------------------------------------------------------------------

  async runSQL(sql: string, rowCap: number): Promise<SparkResult> {
    await this.ensureReady();
    const snippet = callPy("run_sql", sql, rowCap);
    const raw = await this._host.runPython(snippet);
    return parseResult(extractPayload(raw));
  }

  // -------------------------------------------------------------------------
  // schemaOf
  // -------------------------------------------------------------------------

  async schemaOf(sql: string): Promise<ColumnMeta[]> {
    await this.ensureReady();
    const snippet = callPy("schema_of", sql);
    const raw = await this._host.runPython(snippet);
    return parseSchema(extractPayload(raw));
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
      // Clear the connect memo so the next connect() call actually reconnects.
      this._lastConnectUri = null;
      this._lastConnectToken = null;
    }
  }
}
