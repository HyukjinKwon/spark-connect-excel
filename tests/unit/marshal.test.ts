// SPDX-License-Identifier: Apache-2.0
//
// marshal.test.ts — unit tests for src/bridge/marshal.ts.
//
// Covers:
//   - parseResult: valid SparkResult JSON → SparkResult object
//   - parseSchema: valid schema JSON → ColumnMeta[]
//   - parseConnectResult: {ok:true} is a no-op success; {ok:false,error} throws
//   - {ok:false,error} envelope throws with the Python error name + message
//
// No Office.js, no network, no Pyodide.

import { describe, it, expect } from "vitest";
import {
  parseResult,
  parseSchema,
  parseConnectResult,
} from "../../src/bridge/marshal.js";
import type { SparkResult, ColumnMeta } from "../../src/seam.js";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeSparkResult(overrides?: Partial<SparkResult>): SparkResult {
  return {
    schema: [
      { name: "id", type: "bigint" },
      { name: "name", type: "string" },
    ],
    rows: [[1, "alice"], [2, "bob"]],
    rowCount: 2,
    truncated: false,
    ...overrides,
  };
}

function errorEnvelope(name: string, message: string): string {
  return JSON.stringify({ ok: false, error: { name, message } });
}

// ---------------------------------------------------------------------------
// parseResult
// ---------------------------------------------------------------------------

describe("parseResult — success paths", () => {
  it("parses a basic SparkResult", () => {
    const result = makeSparkResult();
    const parsed = parseResult(JSON.stringify(result));
    expect(parsed.schema).toEqual(result.schema);
    expect(parsed.rows).toEqual(result.rows);
    expect(parsed.rowCount).toBe(2);
    expect(parsed.truncated).toBe(false);
  });

  it("parses a truncated result", () => {
    const result = makeSparkResult({ truncated: true, rowCount: 10000 });
    const parsed = parseResult(JSON.stringify(result));
    expect(parsed.truncated).toBe(true);
    expect(parsed.rowCount).toBe(10000);
  });

  it("parses an empty result set (rowCount 0)", () => {
    const result = makeSparkResult({ rows: [], rowCount: 0 });
    const parsed = parseResult(JSON.stringify(result));
    expect(parsed.rows).toEqual([]);
    expect(parsed.rowCount).toBe(0);
    expect(parsed.truncated).toBe(false);
  });

  it("parses a result with null cells (missing values)", () => {
    const result = makeSparkResult({
      schema: [{ name: "val", type: "double" }],
      rows: [[null], [3.14], [null]],
      rowCount: 3,
    });
    const parsed = parseResult(JSON.stringify(result));
    expect(parsed.rows[0]).toEqual([null]);
    expect(parsed.rows[1]).toEqual([3.14]);
  });

  it("parses a result with all supported Spark types in the schema", () => {
    const schema: ColumnMeta[] = [
      { name: "a", type: "bigint" },
      { name: "b", type: "string" },
      { name: "c", type: "double" },
      { name: "d", type: "boolean" },
      { name: "e", type: "timestamp" },
      { name: "f", type: "date" },
      { name: "g", type: "decimal(10,2)" },
    ];
    const result = makeSparkResult({ schema, rows: [], rowCount: 0 });
    const parsed = parseResult(JSON.stringify(result));
    expect(parsed.schema).toHaveLength(7);
    expect(parsed.schema[4]?.type).toBe("timestamp");
  });

  it("round-trips schema names exactly", () => {
    const schema: ColumnMeta[] = [
      { name: "customer_id", type: "bigint" },
      { name: "total_revenue_usd", type: "decimal(18,2)" },
    ];
    const result = makeSparkResult({ schema, rows: [], rowCount: 0 });
    const parsed = parseResult(JSON.stringify(result));
    expect(parsed.schema[0]?.name).toBe("customer_id");
    expect(parsed.schema[1]?.name).toBe("total_revenue_usd");
  });
});

describe("parseResult — error envelope", () => {
  it("throws when ok=false with an AnalysisException", () => {
    const json = errorEnvelope("AnalysisException", "Table not found: missing_table");
    expect(() => parseResult(json)).toThrow("AnalysisException");
    expect(() => parseResult(json)).toThrow("Table not found: missing_table");
  });

  it("throws when ok=false with a RuntimeError", () => {
    const json = errorEnvelope("RuntimeError", "Not connected — call connect() first.");
    expect(() => parseResult(json)).toThrow("RuntimeError");
  });

  it("error message includes both the name and the message", () => {
    const json = errorEnvelope("ValueError", "invalid row cap");
    let thrown: Error | null = null;
    try {
      parseResult(json);
    } catch (e) {
      thrown = e as Error;
    }
    expect(thrown).not.toBeNull();
    expect(thrown?.message).toContain("ValueError");
    expect(thrown?.message).toContain("invalid row cap");
  });

  it("throws an Error instance (not a string or unknown)", () => {
    const json = errorEnvelope("SomeError", "some message");
    expect(() => parseResult(json)).toThrow(Error);
  });
});

// ---------------------------------------------------------------------------
// parseSchema
// ---------------------------------------------------------------------------

describe("parseSchema — success paths", () => {
  it("parses a valid schema envelope", () => {
    const json = JSON.stringify({
      schema: [
        { name: "id", type: "bigint" },
        { name: "label", type: "string" },
      ],
    });
    const cols = parseSchema(json);
    expect(cols).toHaveLength(2);
    expect(cols[0]).toEqual({ name: "id", type: "bigint" });
    expect(cols[1]).toEqual({ name: "label", type: "string" });
  });

  it("parses an empty schema (query with no columns)", () => {
    const json = JSON.stringify({ schema: [] });
    const cols = parseSchema(json);
    expect(cols).toEqual([]);
  });

  it("parses complex types in schema", () => {
    const json = JSON.stringify({
      schema: [
        { name: "tags", type: "array<string>" },
        { name: "meta", type: "map<string,string>" },
        { name: "nested", type: "struct<a:int,b:string>" },
      ],
    });
    const cols = parseSchema(json);
    expect(cols[0]?.type).toBe("array<string>");
    expect(cols[2]?.type).toBe("struct<a:int,b:string>");
  });
});

describe("parseSchema — error envelope", () => {
  it("throws when ok=false", () => {
    const json = errorEnvelope("AnalysisException", "Undefined column");
    expect(() => parseSchema(json)).toThrow("AnalysisException");
    expect(() => parseSchema(json)).toThrow("Undefined column");
  });

  it("includes both name and message in the thrown error", () => {
    const json = errorEnvelope("ParseException", "Syntax error near WHERE");
    let thrown: Error | null = null;
    try {
      parseSchema(json);
    } catch (e) {
      thrown = e as Error;
    }
    expect(thrown?.message).toContain("ParseException");
    expect(thrown?.message).toContain("Syntax error near WHERE");
  });
});

// ---------------------------------------------------------------------------
// parseConnectResult
// ---------------------------------------------------------------------------

describe("parseConnectResult — success", () => {
  it("does not throw for {ok: true}", () => {
    const json = JSON.stringify({ ok: true });
    expect(() => parseConnectResult(json)).not.toThrow();
  });

  it("returns void (undefined) for {ok: true}", () => {
    const json = JSON.stringify({ ok: true });
    const result = parseConnectResult(json);
    expect(result).toBeUndefined();
  });
});

describe("parseConnectResult — error envelope", () => {
  it("throws when ok=false with a connection error", () => {
    const json = errorEnvelope("ConnectionRefusedError", "Connection refused at localhost:8081");
    expect(() => parseConnectResult(json)).toThrow("ConnectionRefusedError");
  });

  it("error message includes both name and message", () => {
    const json = errorEnvelope("SparkConnectException", "Server not available");
    let thrown: Error | null = null;
    try {
      parseConnectResult(json);
    } catch (e) {
      thrown = e as Error;
    }
    expect(thrown?.message).toContain("SparkConnectException");
    expect(thrown?.message).toContain("Server not available");
  });

  it("throws for a non-OK envelope regardless of extra fields", () => {
    const json = JSON.stringify({
      ok: false,
      error: { name: "IOError", message: "timeout" },
      extra: "ignored",
    });
    expect(() => parseConnectResult(json)).toThrow("IOError");
  });
});

// ---------------------------------------------------------------------------
// Edge cases: malformed but non-error payloads are passed through
// ---------------------------------------------------------------------------

describe("parseResult / parseSchema — non-error payloads without ok=false", () => {
  it("parseResult does NOT throw for payloads without an ok field", () => {
    // Python success payloads don't carry an 'ok' key.
    const result = makeSparkResult();
    expect(() => parseResult(JSON.stringify(result))).not.toThrow();
  });

  it("parseSchema does NOT throw for the plain schema envelope", () => {
    const json = JSON.stringify({ schema: [{ name: "x", type: "int" }] });
    expect(() => parseSchema(json)).not.toThrow();
  });
});
