// SPDX-License-Identifier: Apache-2.0
//
// Unit tests for SparkBridgeHost — the dialog-side orchestration of the
// SparkBridge seam. We drive it with a fake RuntimeHost (no Pyodide, no
// browser) that returns canned JSON for the connect/run_sql/schema_of snippets,
// exercising ensureReady idempotency, marshalling, connect-state tracking, and
// error propagation.

import { describe, it, expect, beforeEach } from "vitest";
import { SparkBridgeHost } from "../../src/bridge/sparkBridgeHost";
import type { RuntimeHost, BootOptions, SparkResult } from "../../src/seam";

const SAMPLE_RESULT: SparkResult = {
  schema: [
    { name: "id", type: "bigint" },
    { name: "name", type: "string" },
  ],
  rows: [
    [1, "a"],
    [2, "b"],
  ],
  rowCount: 2,
  truncated: false,
};

/** A RuntimeHost stand-in that records calls and returns canned snippet output. */
class FakeRuntimeHost implements RuntimeHost {
  bootCount = 0;
  runCalls: string[] = [];
  ready = false;
  terminated = false;

  /** Override to simulate a connect() error envelope. */
  connectResponse = JSON.stringify({ ok: true });
  /** Override to simulate a run_sql() error envelope or a different result. */
  runSqlResponse = JSON.stringify(SAMPLE_RESULT);
  schemaResponse = JSON.stringify({ schema: SAMPLE_RESULT.schema });

  async boot(_opts?: BootOptions, onProgress?: (msg: string) => void): Promise<void> {
    this.bootCount += 1;
    onProgress?.("booting");
    this.ready = true;
  }

  async runPython(src: string): Promise<string> {
    this.runCalls.push(src);
    if (src.includes("run_sql(")) return this.runSqlResponse;
    if (src.includes("schema_of(")) return this.schemaResponse;
    if (src.includes("connect(")) return this.connectResponse;
    // install() snippet and the module loader snippet.
    return '"ok"';
  }

  terminate(): void {
    this.terminated = true;
  }
}

describe("SparkBridgeHost — ensureReady", () => {
  let host: FakeRuntimeHost;
  let bridge: SparkBridgeHost;

  beforeEach(() => {
    host = new FakeRuntimeHost();
    bridge = new SparkBridgeHost(host);
  });

  it("boots the runtime exactly once across concurrent and repeated calls", async () => {
    await Promise.all([bridge.ensureReady(), bridge.ensureReady()]);
    await bridge.ensureReady();
    expect(host.bootCount).toBe(1);
  });

  it("installs pyspark-connect-web and loads the runtime module", async () => {
    await bridge.ensureReady();
    const joined = host.runCalls.join("\n");
    expect(joined).toContain("pyspark_connect_web");
    expect(joined).toContain("spark_excel_runtime");
  });

  it("reports pyodideReady only after ensureReady resolves", async () => {
    expect(bridge.status().pyodideReady).toBe(false);
    await bridge.ensureReady();
    expect(bridge.status().pyodideReady).toBe(true);
  });
});

describe("SparkBridgeHost — connect", () => {
  it("marks the session connected on success", async () => {
    const host = new FakeRuntimeHost();
    const bridge = new SparkBridgeHost(host);
    await bridge.connect("sc://localhost:8081/;transport=grpcweb", { token: "secret" });
    expect(bridge.status().connected).toBe(true);
  });

  it("propagates a Python connect error and stays disconnected", async () => {
    const host = new FakeRuntimeHost();
    host.connectResponse = JSON.stringify({
      ok: false,
      error: { name: "SparkConnectGrpcException", message: "UNAVAILABLE" },
    });
    const bridge = new SparkBridgeHost(host);
    await expect(bridge.connect("sc://bad:1/;transport=grpcweb")).rejects.toThrow("UNAVAILABLE");
    expect(bridge.status().connected).toBe(false);
  });

  it("S-3: skips Python connect() when called again with identical args", async () => {
    const host = new FakeRuntimeHost();
    const bridge = new SparkBridgeHost(host);
    const uri = "sc://localhost:8081/;transport=grpcweb";
    await bridge.connect(uri, { token: "tok" });
    const callsAfterFirst = host.runCalls.filter((c) => c.includes("connect(")).length;
    // Call again with same args — should be a no-op
    await bridge.connect(uri, { token: "tok" });
    const callsAfterSecond = host.runCalls.filter((c) => c.includes("connect(")).length;
    expect(callsAfterSecond).toBe(callsAfterFirst);
  });

  it("S-3: re-runs Python connect() when called with different uri", async () => {
    const host = new FakeRuntimeHost();
    const bridge = new SparkBridgeHost(host);
    await bridge.connect("sc://host-a:8081/", { token: "tok" });
    const callsBefore = host.runCalls.filter((c) => c.includes("connect(")).length;
    await bridge.connect("sc://host-b:8081/", { token: "tok" });
    const callsAfter = host.runCalls.filter((c) => c.includes("connect(")).length;
    expect(callsAfter).toBe(callsBefore + 1);
  });

  it("S-3: cancel() clears memo so next connect() actually reconnects", async () => {
    const host = new FakeRuntimeHost();
    const bridge = new SparkBridgeHost(host);
    const uri = "sc://localhost:8081/;transport=grpcweb";
    await bridge.connect(uri, { token: "tok" });
    const callsAfterFirst = host.runCalls.filter((c) => c.includes("connect(")).length;
    bridge.cancel();
    // After cancel the memo is cleared; same args should re-run connect
    await bridge.connect(uri, { token: "tok" });
    const callsAfterReconnect = host.runCalls.filter((c) => c.includes("connect(")).length;
    expect(callsAfterReconnect).toBe(callsAfterFirst + 1);
  });
});

describe("SparkBridgeHost — stdout marker payload (worker contract)", () => {
  // The real worker captures stdout, so callPy PRINTs its JSON bracketed by a
  // unique marker. Simulate that here (with incidental warnings around it) to
  // prove the host recovers the payload and not the noise. This is the exact
  // failure that left the demo stuck at "Unexpected end of JSON input".
  const MARK = "<<<__SCX_PCW__>>>";

  class MarkerHost extends FakeRuntimeHost {
    override async runPython(src: string): Promise<string> {
      this.runCalls.push(src);
      const wrap = (json: string) => `some pyodide warning\n${MARK}${json}${MARK}`;
      if (src.includes("run_sql(")) return wrap(this.runSqlResponse);
      if (src.includes("schema_of(")) return wrap(this.schemaResponse);
      if (src.includes("connect(")) return wrap(this.connectResponse);
      return '"ok"';
    }
  }

  it("recovers connect()/run_sql() JSON from marker-bracketed stdout with noise", async () => {
    const host = new MarkerHost();
    const bridge = new SparkBridgeHost(host);
    await bridge.connect("sc://localhost:8081/;transport=grpcweb");
    expect(bridge.status().connected).toBe(true);
    const result = await bridge.runSQL("SELECT 1", 10);
    expect(result).toEqual(SAMPLE_RESULT);
  });

  it("still surfaces a Python error envelope when marker-wrapped", async () => {
    const host = new MarkerHost();
    host.runSqlResponse = JSON.stringify({
      ok: false,
      error: { name: "AnalysisException", message: "boom" },
    });
    const bridge = new SparkBridgeHost(host);
    await expect(bridge.runSQL("SELECT 1", 10)).rejects.toThrow("AnalysisException: boom");
  });
});

describe("SparkBridgeHost — ensureReady retry after failure (B-4)", () => {
  it("resets _bootPromise on failure so a later call can retry", async () => {
    let failCount = 0;
    const host = new FakeRuntimeHost();
    // Make the first boot() call fail
    host.boot = async () => {
      failCount++;
      if (failCount === 1) {
        throw new Error("boot failed");
      }
      host.ready = true;
    };

    const bridge = new SparkBridgeHost(host);
    await expect(bridge.ensureReady()).rejects.toThrow("boot failed");

    // After the failure the bridge should allow a retry
    await bridge.ensureReady();
    expect(bridge.status().pyodideReady).toBe(true);
  });
});

describe("SparkBridgeHost — runSQL / schemaOf", () => {
  it("returns a marshalled SparkResult", async () => {
    const host = new FakeRuntimeHost();
    const bridge = new SparkBridgeHost(host);
    const result = await bridge.runSQL("SELECT 1", 10);
    expect(result).toEqual(SAMPLE_RESULT);
  });

  it("auto-boots before running (ensureReady is implied)", async () => {
    const host = new FakeRuntimeHost();
    const bridge = new SparkBridgeHost(host);
    await bridge.runSQL("SELECT 1", 10);
    expect(host.bootCount).toBe(1);
  });

  it("rejects with the Python error name and message on failure", async () => {
    const host = new FakeRuntimeHost();
    host.runSqlResponse = JSON.stringify({
      ok: false,
      error: { name: "AnalysisException", message: "Table or view not found: nope" },
    });
    const bridge = new SparkBridgeHost(host);
    await expect(bridge.runSQL("SELECT * FROM nope", 10)).rejects.toThrow(
      "AnalysisException: Table or view not found: nope",
    );
  });

  it("resolves schemaOf to ColumnMeta[]", async () => {
    const host = new FakeRuntimeHost();
    const bridge = new SparkBridgeHost(host);
    const schema = await bridge.schemaOf("SELECT 1");
    expect(schema).toEqual(SAMPLE_RESULT.schema);
  });

  it("passes the row cap through to the run_sql snippet (b64-encoded)", async () => {
    const host = new FakeRuntimeHost();
    const bridge = new SparkBridgeHost(host);
    await bridge.runSQL("SELECT 1", 4242);
    const runSqlCall = host.runCalls.find((c) => c.includes("run_sql("));
    expect(runSqlCall).toBeDefined();
    // Args are base64-encoded; decode the b64 literal from the snippet and
    // verify that the row cap value is present in the decoded JSON payload.
    const b64Match = /b64decode\("([^"]+)"\)/.exec(runSqlCall!);
    expect(b64Match).not.toBeNull();
    const decoded = atob(b64Match![1]);
    expect(decoded).toContain("4242");
  });
});
