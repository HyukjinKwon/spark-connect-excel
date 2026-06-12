// SPDX-License-Identifier: Apache-2.0
//
// seam.ts — THE FROZEN SEAM (integrator-owned).
//
// This is the stable interface between the Excel-facing lanes (E/F/G/H, which
// run in the task pane) and the runtime lanes (B/C/D, which run inside the COI
// dialog window hosting Pyodide + pyspark-connect-web).
//
// Do not change a signature here without a note in COORDINATION.md — every other
// lane builds against these shapes. The two halves of the seam:
//
//   1. `SparkBridge`        — the async API the task pane calls.
//   2. the message envelope — how that API is tunnelled across the Office Dialog
//                             boundary (Office dialog messages are STRINGS only).
//
// The same `SparkBridge` interface is implemented twice:
//   * SparkBridgeHost   (dialog side)   — does the real work via __pcwRunPython.
//   * SparkBridgeClient (task-pane side) — forwards calls over the envelope.

/** A single result column. `type` is the Spark SQL type name (e.g. "bigint", "string", "timestamp"). */
export interface ColumnMeta {
  name: string;
  type: string;
}

/** A query result, already marshalled to JS-native values. */
export interface SparkResult {
  schema: ColumnMeta[];
  /** Row-major; cells are number | string | boolean | null. Dates/timestamps are ISO-8601 strings. */
  rows: unknown[][];
  /** Rows returned after applying the cap. */
  rowCount: number;
  /** True if the row cap clipped the full result. */
  truncated: boolean;
}

/** Snapshot of runtime readiness; cheap, synchronous, safe to poll. */
export interface RuntimeStatus {
  crossOriginIsolated: boolean;
  pyodideReady: boolean;
  connected: boolean;
}

export interface ConnectOptions {
  /** Bearer token forwarded to the Envoy proxy as `Authorization: Bearer <token>`. */
  token?: string;
}

/**
 * The core contract. `runSQL` is `spark.sql(sql).limit(rowCap + 1).toPandas()`
 * marshalled to a SparkResult; the `+1` detects truncation.
 */
export interface SparkBridge {
  /** Resolve once Pyodide is loaded, pcw.install() has run, and COI is verified. Idempotent. */
  ensureReady(): Promise<void>;
  /** Build/refresh the SparkSession against the given grpc-web endpoint. */
  connect(uri: string, opts?: ConnectOptions): Promise<void>;
  /** Run a SQL query and return at most `rowCap` rows (+ truncation flag). */
  runSQL(sql: string, rowCap: number): Promise<SparkResult>;
  /** Resolve the result schema without fetching data (for chart-type inference). */
  schemaOf(sql: string): Promise<ColumnMeta[]>;
  /** Cheap readiness snapshot. */
  status(): RuntimeStatus;
  /** Best-effort cancel of an in-flight query. */
  cancel(): void;
}

// ---------------------------------------------------------------------------
// Runtime host (Lane C provides; Lanes B & D consume via this interface, never
// the concrete class — keeps the dialog/bridge decoupled from Pyodide wiring).
// ---------------------------------------------------------------------------

export interface BootOptions {
  /** Override the Pyodide CDN index URL. */
  pyodideIndexUrl?: string;
  /** URL of the pyspark-connect-web wheel (or a PyPI spec micropip can resolve). */
  wheelUrl?: string;
}

export interface RuntimeHost {
  /** Boot Pyodide + install the wheel. Idempotent; reports human-readable progress. */
  boot(opts?: BootOptions, onProgress?: (msg: string) => void): Promise<void>;
  /** Execute Python source in the worker, resolving with the stringified last expression. */
  runPython(src: string): Promise<string>;
  /** True once boot() has completed. */
  readonly ready: boolean;
  /** Tear down the worker. */
  terminate(): void;
}

// ---------------------------------------------------------------------------
// Message envelope (task pane <-> dialog). All payloads are JSON strings.
// ---------------------------------------------------------------------------

export type BridgeMethod = "ensureReady" | "connect" | "runSQL" | "schemaOf" | "cancel";

export interface BridgeRequest {
  kind: "req";
  id: number;
  method: BridgeMethod;
  args: unknown[];
}

export interface BridgeResponse {
  kind: "res";
  id: number;
  ok: boolean;
  result?: unknown;
  error?: { name: string; message: string };
}

/** Unsolicited dialog -> parent pushes (boot progress, status changes, logs). */
export interface BridgeEvent {
  kind: "evt";
  /**
   * `unsupported` signals the host can't run the engine (no cross-origin
   * isolation / SharedArrayBuffer) — the task pane shows a blocking guidance
   * panel instead of a generic error. payload: `{ reason: string }`.
   */
  event: "ready" | "status" | "progress" | "log" | "unsupported";
  payload?: unknown;
}

export type BridgeMessage = BridgeRequest | BridgeResponse | BridgeEvent;

export function encodeMessage(m: BridgeMessage): string {
  return JSON.stringify(m);
}

export function decodeMessage(s: string): BridgeMessage {
  return JSON.parse(s) as BridgeMessage;
}

/** Canonical grpc-web connection URI scheme, matching pyspark-connect-web. */
export function normalizeRemoteUri(host: string, port: number, tls: boolean): string {
  const scheme = tls ? "https" : "http";
  // pyspark-connect-web accepts both the sc:// form and the http(s):// shorthand.
  return `${scheme}://${host}:${port}`;
}

export const SC_URI_HINT = "sc://host:port/;transport=grpcweb";
