// SPDX-License-Identifier: Apache-2.0
//
// seam.test.ts — unit tests for the pure helpers exported from src/seam.ts.
//
// Tests:
//   - encodeMessage / decodeMessage round-trip for req / res / evt variants
//   - normalizeRemoteUri output shape
//   - SC_URI_HINT is the sc:// form
//
// No Office.js, no network, no Pyodide.

import { describe, it, expect } from "vitest";
import {
  encodeMessage,
  decodeMessage,
  normalizeRemoteUri,
  SC_URI_HINT,
  type BridgeRequest,
  type BridgeResponse,
  type BridgeEvent,
} from "../../src/seam.js";

// ---------------------------------------------------------------------------
// encodeMessage / decodeMessage round-trips
// ---------------------------------------------------------------------------

describe("encodeMessage / decodeMessage — req round-trip", () => {
  it("round-trips a BridgeRequest with no args", () => {
    const req: BridgeRequest = {
      kind: "req",
      id: 1,
      method: "ensureReady",
      args: [],
    };
    const encoded = encodeMessage(req);
    expect(typeof encoded).toBe("string");
    const decoded = decodeMessage(encoded) as BridgeRequest;
    expect(decoded.kind).toBe("req");
    expect(decoded.id).toBe(1);
    expect(decoded.method).toBe("ensureReady");
    expect(decoded.args).toEqual([]);
  });

  it("round-trips a BridgeRequest with args", () => {
    const req: BridgeRequest = {
      kind: "req",
      id: 42,
      method: "runSQL",
      args: ["SELECT 1", 1000],
    };
    const decoded = decodeMessage(encodeMessage(req)) as BridgeRequest;
    expect(decoded.kind).toBe("req");
    expect(decoded.id).toBe(42);
    expect(decoded.method).toBe("runSQL");
    expect(decoded.args).toEqual(["SELECT 1", 1000]);
  });

  it("round-trips a BridgeRequest for connect with uri and opts", () => {
    const req: BridgeRequest = {
      kind: "req",
      id: 7,
      method: "connect",
      args: ["sc://localhost:8081/;transport=grpcweb", { token: "tok123" }],
    };
    const decoded = decodeMessage(encodeMessage(req)) as BridgeRequest;
    expect(decoded.args[0]).toBe("sc://localhost:8081/;transport=grpcweb");
    expect((decoded.args[1] as Record<string, unknown>)?.token).toBe("tok123");
  });

  it("round-trips all BridgeMethod values", () => {
    const methods: BridgeRequest["method"][] = [
      "ensureReady",
      "connect",
      "runSQL",
      "schemaOf",
      "cancel",
    ];
    for (const method of methods) {
      const req: BridgeRequest = { kind: "req", id: 1, method, args: [] };
      const decoded = decodeMessage(encodeMessage(req)) as BridgeRequest;
      expect(decoded.method).toBe(method);
    }
  });
});

describe("encodeMessage / decodeMessage — res round-trip", () => {
  it("round-trips a successful BridgeResponse with a SparkResult payload", () => {
    const res: BridgeResponse = {
      kind: "res",
      id: 5,
      ok: true,
      result: {
        schema: [{ name: "id", type: "bigint" }],
        rows: [[1], [2]],
        rowCount: 2,
        truncated: false,
      },
    };
    const decoded = decodeMessage(encodeMessage(res)) as BridgeResponse;
    expect(decoded.kind).toBe("res");
    expect(decoded.id).toBe(5);
    expect(decoded.ok).toBe(true);
    expect(decoded.result).toEqual(res.result);
  });

  it("round-trips a failed BridgeResponse with error info", () => {
    const res: BridgeResponse = {
      kind: "res",
      id: 9,
      ok: false,
      error: { name: "RuntimeError", message: "Not connected" },
    };
    const decoded = decodeMessage(encodeMessage(res)) as BridgeResponse;
    expect(decoded.kind).toBe("res");
    expect(decoded.ok).toBe(false);
    expect(decoded.error?.name).toBe("RuntimeError");
    expect(decoded.error?.message).toBe("Not connected");
  });

  it("round-trips a successful BridgeResponse with no result (void)", () => {
    const res: BridgeResponse = { kind: "res", id: 3, ok: true };
    const decoded = decodeMessage(encodeMessage(res)) as BridgeResponse;
    expect(decoded.kind).toBe("res");
    expect(decoded.ok).toBe(true);
    expect(decoded.result).toBeUndefined();
  });
});

describe("encodeMessage / decodeMessage — evt round-trip", () => {
  it("round-trips a 'ready' BridgeEvent", () => {
    const evt: BridgeEvent = { kind: "evt", event: "ready" };
    const decoded = decodeMessage(encodeMessage(evt)) as BridgeEvent;
    expect(decoded.kind).toBe("evt");
    expect(decoded.event).toBe("ready");
    expect(decoded.payload).toBeUndefined();
  });

  it("round-trips a 'status' BridgeEvent with payload", () => {
    const evt: BridgeEvent = {
      kind: "evt",
      event: "status",
      payload: { crossOriginIsolated: true, pyodideReady: true, connected: false },
    };
    const decoded = decodeMessage(encodeMessage(evt)) as BridgeEvent;
    expect(decoded.kind).toBe("evt");
    expect(decoded.event).toBe("status");
    expect(decoded.payload).toEqual(evt.payload);
  });

  it("round-trips a 'progress' BridgeEvent", () => {
    const evt: BridgeEvent = {
      kind: "evt",
      event: "progress",
      payload: "Loading Pyodide…",
    };
    const decoded = decodeMessage(encodeMessage(evt)) as BridgeEvent;
    expect(decoded.event).toBe("progress");
    expect(decoded.payload).toBe("Loading Pyodide…");
  });

  it("round-trips a 'log' BridgeEvent", () => {
    const evt: BridgeEvent = {
      kind: "evt",
      event: "log",
      payload: "pyspark-connect-web installed",
    };
    const decoded = decodeMessage(encodeMessage(evt)) as BridgeEvent;
    expect(decoded.event).toBe("log");
  });
});

describe("encodeMessage produces valid JSON strings", () => {
  it("encode output is parseable by JSON.parse", () => {
    const req: BridgeRequest = { kind: "req", id: 1, method: "cancel", args: [] };
    const encoded = encodeMessage(req);
    expect(() => JSON.parse(encoded)).not.toThrow();
  });

  it("encodes special characters in SQL args without losing them", () => {
    const req: BridgeRequest = {
      kind: "req",
      id: 1,
      method: "runSQL",
      args: ["SELECT '日本語' AS text, \"col\" FROM t WHERE x > 0", 100],
    };
    const decoded = decodeMessage(encodeMessage(req)) as BridgeRequest;
    expect(decoded.args[0]).toBe("SELECT '日本語' AS text, \"col\" FROM t WHERE x > 0");
  });
});

// ---------------------------------------------------------------------------
// normalizeRemoteUri
// ---------------------------------------------------------------------------

describe("normalizeRemoteUri", () => {
  it("produces an http:// URI for non-TLS", () => {
    const uri = normalizeRemoteUri("localhost", 8081, false);
    expect(uri).toBe("http://localhost:8081");
  });

  it("produces an https:// URI for TLS", () => {
    const uri = normalizeRemoteUri("spark.example.com", 8443, true);
    expect(uri).toBe("https://spark.example.com:8443");
  });

  it("includes the port in all cases", () => {
    expect(normalizeRemoteUri("host", 3000, false)).toContain(":3000");
    expect(normalizeRemoteUri("host", 443, true)).toContain(":443");
  });

  it("uses the host exactly as given", () => {
    const uri = normalizeRemoteUri("my-cluster.internal", 8081, false);
    expect(uri).toContain("my-cluster.internal");
  });
});

// ---------------------------------------------------------------------------
// SC_URI_HINT — must be the sc:// form
// ---------------------------------------------------------------------------

describe("SC_URI_HINT", () => {
  it("is a string", () => {
    expect(typeof SC_URI_HINT).toBe("string");
  });

  it("starts with sc://", () => {
    expect(SC_URI_HINT.startsWith("sc://")).toBe(true);
  });

  it("contains transport=grpcweb", () => {
    expect(SC_URI_HINT).toContain("transport=grpcweb");
  });

  it("matches the documented example form sc://host:port/;transport=grpcweb", () => {
    // Must look like sc://<host:port>/;transport=grpcweb (placeholders, not digits)
    expect(SC_URI_HINT).toMatch(/^sc:\/\/[^/]+\/;transport=grpcweb$/);
  });
});
