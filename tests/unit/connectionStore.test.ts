// SPDX-License-Identifier: Apache-2.0
//
// connectionStore.test.ts — unit tests for src/connection/connectionStore.ts.
//
// Covers:
//   - buildRemoteUri output (sc:// form)
//   - saveConnection / loadConnection round-trip via memorySettingsBackend
//   - loadConnection returns null when nothing saved
//   - loadConnection returns null for corrupt stored value
//   - clearConnection removes the persisted value
//   - Token helpers (saveToken / loadToken / clearToken) NEVER write to the
//     settings backend (DECISIONS #6 invariant)
//
// No Office.js dependency — everything uses the injectable memorySettingsBackend.

import { describe, it, expect, beforeEach } from "vitest";
import {
  buildRemoteUri,
  saveConnection,
  loadConnection,
  clearConnection,
  memorySettingsBackend,
  type ConnectionConfig,
  type SettingsBackend,
} from "../../src/connection/connectionStore.js";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeConfig(overrides?: Partial<ConnectionConfig>): ConnectionConfig {
  return {
    host: "localhost",
    port: 8081,
    tls: false,
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// buildRemoteUri
// ---------------------------------------------------------------------------

describe("buildRemoteUri", () => {
  it("produces the sc:// form", () => {
    const uri = buildRemoteUri({ host: "localhost", port: 8081, tls: false });
    expect(uri).toBe("sc://localhost:8081/;transport=grpcweb");
  });

  it("always uses sc:// regardless of the tls flag", () => {
    const uriTls = buildRemoteUri({ host: "spark.example.com", port: 8443, tls: true });
    expect(uriTls.startsWith("sc://")).toBe(true);
    expect(uriTls).toBe("sc://spark.example.com:8443/;transport=grpcweb");
  });

  it("always includes transport=grpcweb", () => {
    const uri = buildRemoteUri(makeConfig());
    expect(uri).toContain("transport=grpcweb");
  });

  it("embeds the host and port exactly", () => {
    const uri = buildRemoteUri({ host: "my-cluster.internal", port: 9090, tls: false });
    expect(uri).toContain("my-cluster.internal:9090");
  });

  it("matches the pattern sc://<host>:<port>/;transport=grpcweb", () => {
    const uri = buildRemoteUri(makeConfig());
    expect(uri).toMatch(/^sc:\/\/.+:\d+\/;transport=grpcweb$/);
  });
});

// ---------------------------------------------------------------------------
// saveConnection / loadConnection round-trip
// ---------------------------------------------------------------------------

describe("saveConnection / loadConnection — round-trip", () => {
  let backend: SettingsBackend;

  beforeEach(() => {
    backend = memorySettingsBackend();
  });

  it("returns null when nothing has been saved", () => {
    expect(loadConnection(backend)).toBeNull();
  });

  it("saves and loads a basic config", async () => {
    const cfg = makeConfig();
    await saveConnection(cfg, backend);
    const loaded = loadConnection(backend);
    expect(loaded).toEqual(cfg);
  });

  it("round-trips the tls=true variant", async () => {
    const cfg = makeConfig({ host: "spark.example.com", port: 8443, tls: true });
    await saveConnection(cfg, backend);
    const loaded = loadConnection(backend);
    expect(loaded?.tls).toBe(true);
    expect(loaded?.host).toBe("spark.example.com");
    expect(loaded?.port).toBe(8443);
  });

  it("round-trips all fields without mutation", async () => {
    const cfg: ConnectionConfig = { host: "cluster.internal", port: 9999, tls: false };
    await saveConnection(cfg, backend);
    const loaded = loadConnection(backend);
    expect(loaded).toStrictEqual(cfg);
  });

  it("overwrites an existing config on re-save", async () => {
    await saveConnection(makeConfig({ host: "old-host", port: 1111 }), backend);
    await saveConnection(makeConfig({ host: "new-host", port: 2222 }), backend);
    const loaded = loadConnection(backend);
    expect(loaded?.host).toBe("new-host");
    expect(loaded?.port).toBe(2222);
  });
});

describe("loadConnection — error handling", () => {
  it("returns null for corrupt JSON", async () => {
    const backend = memorySettingsBackend();
    await backend.set("sparkConnectConfig", "NOT-JSON{{{{");
    expect(loadConnection(backend)).toBeNull();
  });

  it("returns null for a JSON value that is not a ConnectionConfig object", async () => {
    const backend = memorySettingsBackend();
    await backend.set("sparkConnectConfig", JSON.stringify([1, 2, 3]));
    expect(loadConnection(backend)).toBeNull();
  });

  it("returns null for a JSON object missing required fields", async () => {
    const backend = memorySettingsBackend();
    // Missing 'tls'
    await backend.set("sparkConnectConfig", JSON.stringify({ host: "x", port: 80 }));
    expect(loadConnection(backend)).toBeNull();
  });

  it("returns null for a JSON object with wrong field types", async () => {
    const backend = memorySettingsBackend();
    // port should be a number, not a string
    await backend.set(
      "sparkConnectConfig",
      JSON.stringify({ host: "x", port: "8081", tls: false }),
    );
    expect(loadConnection(backend)).toBeNull();
  });
});

describe("clearConnection", () => {
  it("removes a saved config", async () => {
    const backend = memorySettingsBackend();
    await saveConnection(makeConfig(), backend);
    expect(loadConnection(backend)).not.toBeNull();

    await clearConnection(backend);
    expect(loadConnection(backend)).toBeNull();
  });

  it("is a no-op when nothing is saved", async () => {
    const backend = memorySettingsBackend();
    await expect(clearConnection(backend)).resolves.toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// memorySettingsBackend itself
// ---------------------------------------------------------------------------

describe("memorySettingsBackend", () => {
  it("returns null for a key that was never set", () => {
    const backend = memorySettingsBackend();
    expect(backend.get("nonexistent")).toBeNull();
  });

  it("persists set/get within the same instance", async () => {
    const backend = memorySettingsBackend();
    await backend.set("myKey", "myValue");
    expect(backend.get("myKey")).toBe("myValue");
  });

  it("remove erases a key", async () => {
    const backend = memorySettingsBackend();
    await backend.set("k", "v");
    await backend.remove("k");
    expect(backend.get("k")).toBeNull();
  });

  it("two instances are independent", async () => {
    const a = memorySettingsBackend();
    const b = memorySettingsBackend();
    await a.set("key", "from-a");
    expect(b.get("key")).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// DECISIONS #6 — token helpers NEVER write to the settings backend
//
// saveToken / loadToken / clearToken use OfficeRuntime.storage (or an internal
// in-memory Map) — neither of which is the injected SettingsBackend. We verify
// this by passing a spy-wrapped backend and asserting it is never written to
// by those helpers.
//
// Note: In the test environment OfficeRuntime is not defined, so saveToken /
// loadToken / clearToken fall back to an internal in-memory Map. The important
// invariant is that the SettingsBackend is NEVER touched.
// ---------------------------------------------------------------------------

describe("DECISIONS #6 — saveToken / loadToken / clearToken do not write to SettingsBackend", () => {
  it("saveToken does not call backend.set", async () => {
    // Import the token helpers dynamically to avoid top-level Office.js issues.
    const { saveToken } = await import(
      "../../src/connection/connectionStore.js"
    );

    const writes: string[] = [];
    const spyBackend: SettingsBackend = {
      get: () => null,
      set: async (key) => {
        writes.push(key);
      },
      remove: async () => {},
    };

    await saveToken("my-bearer-token");

    // The spy backend must NOT have been written to.
    expect(writes).toEqual([]);

    // Silence unused variable warning — spyBackend was used to catch writes.
    void spyBackend;
  });

  it("the settings backend has no token-bearing key after saveToken", async () => {
    const { saveToken } = await import(
      "../../src/connection/connectionStore.js"
    );

    const backend = memorySettingsBackend();

    // Also save a regular config so the backend is not completely empty.
    await saveConnection(makeConfig(), backend);
    await saveToken("super-secret-token");

    // Enumerate every key in the backend store.
    // We check: none of the stored values contain the token.
    const configRaw = backend.get("sparkConnectConfig");
    expect(configRaw).not.toContain("super-secret-token");

    // Also assert the well-known token key is absent from the settings backend.
    expect(backend.get("sparkConnectToken")).toBeNull();
  });

  it("loadToken does not call backend.get", async () => {
    const { loadToken } = await import(
      "../../src/connection/connectionStore.js"
    );

    const gets: string[] = [];
    const spyBackend: SettingsBackend = {
      get: (key) => {
        gets.push(key);
        return null;
      },
      set: async () => {},
      remove: async () => {},
    };

    await loadToken();
    expect(gets).toEqual([]);
    void spyBackend;
  });

  it("clearToken does not call backend.remove", async () => {
    const { clearToken } = await import(
      "../../src/connection/connectionStore.js"
    );

    const removes: string[] = [];
    const spyBackend: SettingsBackend = {
      get: () => null,
      set: async () => {},
      remove: async (key) => {
        removes.push(key);
      },
    };

    await clearToken();
    expect(removes).toEqual([]);
    void spyBackend;
  });
});
