// SPDX-License-Identifier: Apache-2.0
//
// sparkBridgeClient.test.ts — unit tests for SparkBridgeClient
//   (src/bridge/sparkBridgeClient.ts).
//
// Strategy: inject a FAKE transport ({send, subscribe}) so the client never
// imports or calls Office.js. Tests drive the wire protocol manually.
//
// Covers:
//   - A runSQL call emits a BridgeRequest over the transport
//   - Feeding back a matching BridgeResponse resolves the promise with the result
//   - An error BridgeResponse rejects with the error message
//   - status() is synchronous and reflects pushed BridgeEvent "status" payloads
//   - cancel() sends a BridgeRequest (fire-and-forget, does not wait)
//   - rejectAll() rejects all pending calls

import { describe, it, expect, vi } from "vitest";
import {
  SparkBridgeClient,
  type BridgeTransport,
} from "../../src/bridge/sparkBridgeClient.js";
import {
  decodeMessage,
  encodeMessage,
  type BridgeRequest,
  type BridgeResponse,
  type BridgeEvent,
} from "../../src/seam.js";
import type { SparkResult } from "../../src/seam.js";

// ---------------------------------------------------------------------------
// Fake transport factory
// ---------------------------------------------------------------------------

interface FakeTransport extends BridgeTransport {
  /** All messages the client has sent (decoded). */
  sent: BridgeRequest[];
  /** Push a raw string to all subscribed listeners (simulates the dialog sending). */
  push(raw: string): void;
}

function makeFakeTransport(): FakeTransport {
  const listeners: Array<(msg: string) => void> = [];
  const sent: BridgeRequest[] = [];

  return {
    sent,
    send(msg: string) {
      const decoded = decodeMessage(msg) as BridgeRequest;
      sent.push(decoded);
    },
    subscribe(listener: (msg: string) => void) {
      listeners.push(listener);
      return () => {
        const idx = listeners.indexOf(listener);
        if (idx !== -1) listeners.splice(idx, 1);
      };
    },
    push(raw: string) {
      for (const l of listeners) l(raw);
    },
  };
}

// ---------------------------------------------------------------------------
// Helper: send a matching BridgeResponse back for the last sent request
// ---------------------------------------------------------------------------

function replyOk(transport: FakeTransport, result: unknown): void {
  const last = transport.sent[transport.sent.length - 1]!;
  const res: BridgeResponse = { kind: "res", id: last.id, ok: true, result };
  transport.push(encodeMessage(res));
}

function replyError(transport: FakeTransport, name: string, message: string): void {
  const last = transport.sent[transport.sent.length - 1]!;
  const res: BridgeResponse = {
    kind: "res",
    id: last.id,
    ok: false,
    error: { name, message },
  };
  transport.push(encodeMessage(res));
}

// ---------------------------------------------------------------------------
// BridgeRequest emission
// ---------------------------------------------------------------------------

describe("SparkBridgeClient — emits correct BridgeRequest", () => {
  it("runSQL emits a req with method='runSQL' and correct args", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const promise = client.runSQL("SELECT 1", 500);
    // Resolve it so the test doesn't hang.
    replyOk(transport, {
      schema: [],
      rows: [],
      rowCount: 0,
      truncated: false,
    } satisfies SparkResult);
    await promise;

    expect(transport.sent).toHaveLength(1);
    const req = transport.sent[0]!;
    expect(req.kind).toBe("req");
    expect(req.method).toBe("runSQL");
    expect(req.args[0]).toBe("SELECT 1");
    expect(req.args[1]).toBe(500);
  });

  it("connect emits a req with method='connect' and uri arg", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const promise = client.connect("sc://localhost:8081/;transport=grpcweb");
    replyOk(transport, undefined);
    await promise;

    const req = transport.sent[0]!;
    expect(req.method).toBe("connect");
    expect(req.args[0]).toBe("sc://localhost:8081/;transport=grpcweb");
  });

  it("ensureReady emits a req with method='ensureReady'", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const promise = client.ensureReady();
    replyOk(transport, undefined);
    await promise;

    expect(transport.sent[0]?.method).toBe("ensureReady");
    expect(transport.sent[0]?.args).toEqual([]);
  });

  it("schemaOf emits a req with method='schemaOf' and sql arg", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const promise = client.schemaOf("SELECT id FROM t");
    replyOk(transport, [{ name: "id", type: "bigint" }]);
    await promise;

    expect(transport.sent[0]?.method).toBe("schemaOf");
    expect(transport.sent[0]?.args[0]).toBe("SELECT id FROM t");
  });

  it("each call gets a unique, incrementing id", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const p1 = client.ensureReady();
    replyOk(transport, undefined);
    await p1;

    const p2 = client.ensureReady();
    replyOk(transport, undefined);
    await p2;

    expect(transport.sent[0]?.id).not.toBe(transport.sent[1]?.id);
    expect(transport.sent[1]!.id).toBeGreaterThan(transport.sent[0]!.id);
  });
});

// ---------------------------------------------------------------------------
// BridgeResponse → promise resolution
// ---------------------------------------------------------------------------

describe("SparkBridgeClient — resolves promise from matching BridgeResponse", () => {
  it("runSQL resolves with the SparkResult from the response", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const expected: SparkResult = {
      schema: [{ name: "n", type: "bigint" }],
      rows: [[1], [2], [3]],
      rowCount: 3,
      truncated: false,
    };

    const promise = client.runSQL("SELECT n FROM t", 100);
    replyOk(transport, expected);
    const result = await promise;

    expect(result).toEqual(expected);
  });

  it("connect resolves (void) on a successful response", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const promise = client.connect("sc://host:8081/;transport=grpcweb", {
      token: "tok",
    });
    replyOk(transport, undefined);
    await expect(promise).resolves.toBeUndefined();
  });

  it("ensureReady resolves (void) on success", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const promise = client.ensureReady();
    replyOk(transport, undefined);
    await expect(promise).resolves.toBeUndefined();
  });

  it("schemaOf resolves with the ColumnMeta array", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const cols = [
      { name: "a", type: "bigint" },
      { name: "b", type: "string" },
    ];
    const promise = client.schemaOf("SELECT a,b FROM t");
    replyOk(transport, cols);
    const result = await promise;
    expect(result).toEqual(cols);
  });

  it("only resolves the matching promise by id (in-flight multiplexing)", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    // Start two calls simultaneously.
    const p1 = client.runSQL("SELECT 1", 10);
    const p2 = client.runSQL("SELECT 2", 20);

    const req1 = transport.sent[0]!;
    const req2 = transport.sent[1]!;
    expect(req1.id).not.toBe(req2.id);

    // Reply to p2 first.
    const res2: BridgeResponse = {
      kind: "res",
      id: req2.id,
      ok: true,
      result: { schema: [], rows: [["two"]], rowCount: 1, truncated: false },
    };
    transport.push(encodeMessage(res2));

    // p2 should resolve; p1 is still pending.
    const r2 = await p2;
    expect((r2 as SparkResult).rows[0]).toEqual(["two"]);

    // Now reply to p1.
    const res1: BridgeResponse = {
      kind: "res",
      id: req1.id,
      ok: true,
      result: { schema: [], rows: [["one"]], rowCount: 1, truncated: false },
    };
    transport.push(encodeMessage(res1));
    const r1 = await p1;
    expect((r1 as SparkResult).rows[0]).toEqual(["one"]);
  });
});

// ---------------------------------------------------------------------------
// Error responses → promise rejection
// ---------------------------------------------------------------------------

describe("SparkBridgeClient — rejects promise from error BridgeResponse", () => {
  it("rejects with an Error that includes the name and message", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const promise = client.runSQL("SELECT bad", 10);
    replyError(transport, "AnalysisException", "Table not found");
    await expect(promise).rejects.toThrow("AnalysisException");

    const promise2 = client.runSQL("SELECT bad2", 10);
    replyError(transport, "AnalysisException", "Table not found 2");
    await expect(promise2).rejects.toBeInstanceOf(Error);
  });

  it("rejects with 'Unknown bridge error' when error field is absent", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const promise = client.ensureReady();
    const req = transport.sent[0]!;
    // Omit the error field.
    const res: BridgeResponse = { kind: "res", id: req.id, ok: false };
    transport.push(encodeMessage(res));
    await expect(promise).rejects.toThrow("Unknown bridge error");
  });

  it("rejects when the transport send throws synchronously", async () => {
    const brokenTransport: BridgeTransport = {
      send: () => {
        throw new Error("transport unavailable");
      },
      subscribe: (l) => {
        void l;
        return () => {};
      },
    };
    const client = new SparkBridgeClient(brokenTransport);
    await expect(client.ensureReady()).rejects.toThrow("transport unavailable");
  });
});

// ---------------------------------------------------------------------------
// status() — synchronous, reflects BridgeEvent "status" pushes
// ---------------------------------------------------------------------------

describe("SparkBridgeClient — status()", () => {
  it("returns the default status before any events", () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);
    const s = client.status();
    expect(s.crossOriginIsolated).toBe(false);
    expect(s.pyodideReady).toBe(false);
    expect(s.connected).toBe(false);
  });

  it("updates crossOriginIsolated when a status event is pushed", () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const evt: BridgeEvent = {
      kind: "evt",
      event: "status",
      payload: { crossOriginIsolated: true },
    };
    transport.push(encodeMessage(evt));

    expect(client.status().crossOriginIsolated).toBe(true);
    // Other fields unchanged.
    expect(client.status().pyodideReady).toBe(false);
  });

  it("merges partial status updates without resetting other fields", () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    // First event: pyodideReady
    transport.push(
      encodeMessage({
        kind: "evt",
        event: "status",
        payload: { pyodideReady: true },
      } satisfies BridgeEvent),
    );
    // Second event: connected
    transport.push(
      encodeMessage({
        kind: "evt",
        event: "status",
        payload: { connected: true },
      } satisfies BridgeEvent),
    );

    const s = client.status();
    expect(s.pyodideReady).toBe(true);
    expect(s.connected).toBe(true);
    expect(s.crossOriginIsolated).toBe(false); // not yet set
  });

  it("status() is synchronous (does not return a Promise)", () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);
    const s = client.status();
    expect(s).not.toBeInstanceOf(Promise);
    expect(typeof s.crossOriginIsolated).toBe("boolean");
  });

  it("returns a snapshot copy (not a live reference)", () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const snap1 = client.status();
    transport.push(
      encodeMessage({
        kind: "evt",
        event: "status",
        payload: { pyodideReady: true },
      } satisfies BridgeEvent),
    );
    const snap2 = client.status();

    // snap1 was captured before the event; it should not have changed.
    expect(snap1.pyodideReady).toBe(false);
    expect(snap2.pyodideReady).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// cancel() — fire-and-forget
// ---------------------------------------------------------------------------

describe("SparkBridgeClient — cancel()", () => {
  it("sends a BridgeRequest with method='cancel'", () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    client.cancel();

    expect(transport.sent).toHaveLength(1);
    expect(transport.sent[0]?.method).toBe("cancel");
    expect(transport.sent[0]?.kind).toBe("req");
  });

  it("returns void immediately (no promise)", () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);
    const result = client.cancel();
    expect(result).toBeUndefined();
  });

  it("does not leave a pending entry waiting for a response", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    client.cancel();

    // Start a real runSQL call after cancel; if cancel had leaked a pending
    // entry the id accounting might be off — verify the next call still works.
    const p = client.runSQL("SELECT 1", 10);
    replyOk(transport, { schema: [], rows: [], rowCount: 0, truncated: false });
    await expect(p).resolves.toBeTruthy();
  });
});

// ---------------------------------------------------------------------------
// rejectAll()
// ---------------------------------------------------------------------------

describe("SparkBridgeClient — rejectAll()", () => {
  it("rejects all pending calls with the given reason", async () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);

    const p1 = client.ensureReady();
    const p2 = client.runSQL("SELECT 1", 10);

    client.rejectAll("Dialog closed (code 12006)");

    await expect(p1).rejects.toThrow("Dialog closed (code 12006)");
    await expect(p2).rejects.toThrow("Dialog closed (code 12006)");
  });

  it("is a no-op when there are no pending calls", () => {
    const transport = makeFakeTransport();
    const client = new SparkBridgeClient(transport);
    expect(() => client.rejectAll("reason")).not.toThrow();
  });
});

// ---------------------------------------------------------------------------
// onEvent callback
// ---------------------------------------------------------------------------

describe("SparkBridgeClient — onEvent callback", () => {
  it("invokes onEvent for each BridgeEvent received", () => {
    const transport = makeFakeTransport();
    const received: BridgeEvent[] = [];
    const client = new SparkBridgeClient(transport, {
      onEvent: (evt) => received.push(evt),
    });
    void client; // constructed for its subscribe side-effect

    const evt1: BridgeEvent = { kind: "evt", event: "progress", payload: "Loading…" };
    const evt2: BridgeEvent = { kind: "evt", event: "ready" };

    transport.push(encodeMessage(evt1));
    transport.push(encodeMessage(evt2));

    expect(received).toHaveLength(2);
    expect(received[0]?.event).toBe("progress");
    expect(received[1]?.event).toBe("ready");
  });

  it("does not invoke onEvent for BridgeResponse messages", async () => {
    const transport = makeFakeTransport();
    const onEvent = vi.fn();
    const client = new SparkBridgeClient(transport, { onEvent });

    const p = client.ensureReady();
    replyOk(transport, undefined);
    await p;

    expect(onEvent).not.toHaveBeenCalled();
  });
});
