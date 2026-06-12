// SPDX-License-Identifier: Apache-2.0
//
// marshal.ts — pure helpers for parsing JSON returned by the Python runtime.
//
// The Python runtime returns either a SparkResult-shaped JSON object (success)
// or a {"ok": false, "error": {name, message}} envelope (failure).  All
// "is-this-an-error?" logic lives here so both SparkBridgeHost and tests share
// one implementation.

import type { ColumnMeta, SparkResult } from "../seam";

// ---------------------------------------------------------------------------
// Internal types
// ---------------------------------------------------------------------------

interface ErrorEnvelope {
  ok: false;
  error: { name: string; message: string };
}

/**
 * Detect whether the parsed JSON represents a Python-runtime error envelope.
 *
 * The Python runtime signals failure with `{"ok": false, "error": {...}}`.
 * Successful payloads never carry an `ok` key (SparkResult) or carry
 * `{"ok": true}` (connect response).
 */
function isErrorEnvelope(v: unknown): v is ErrorEnvelope {
  return (
    typeof v === "object" &&
    v !== null &&
    "ok" in v &&
    (v as Record<string, unknown>)["ok"] === false
  );
}

/** Throw a descriptive Error from an ErrorEnvelope. */
function throwEnvelope(env: ErrorEnvelope): never {
  throw new Error(`${env.error.name}: ${env.error.message}`);
}

// ---------------------------------------------------------------------------
// Public helpers
// ---------------------------------------------------------------------------

/**
 * Parse the JSON string returned by the Python `connect()` function.
 *
 * Throws if the payload carries `{ok: false}`.
 */
export function parseConnectResult(json: string): void {
  const parsed = JSON.parse(json) as unknown;
  if (isErrorEnvelope(parsed)) {
    throwEnvelope(parsed);
  }
  // {"ok": true} — success; nothing to return.
}

/**
 * Parse the JSON string returned by the Python `run_sql()` function.
 *
 * On success `run_sql` returns a plain SparkResult object (no `ok` wrapper).
 * On failure it returns `{"ok": false, "error": {...}}`.
 *
 * Throws an Error (with Python exception name + message) on failure.
 */
export function parseResult(json: string): SparkResult {
  const parsed = JSON.parse(json) as unknown;
  if (isErrorEnvelope(parsed)) {
    throwEnvelope(parsed);
  }
  // Trust the Python runtime's shape; a real app would validate more.
  return parsed as SparkResult;
}

/**
 * Parse the JSON string returned by the Python `schema_of()` function.
 *
 * On success returns `ColumnMeta[]`.
 * On failure throws.
 */
export function parseSchema(json: string): ColumnMeta[] {
  const parsed = JSON.parse(json) as unknown;
  if (isErrorEnvelope(parsed)) {
    throwEnvelope(parsed);
  }
  const envelope = parsed as { schema: ColumnMeta[] };
  return envelope.schema;
}
