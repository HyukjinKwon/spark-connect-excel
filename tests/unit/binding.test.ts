// SPDX-License-Identifier: Apache-2.0
//
// binding.test.ts — unit tests for src/excel/binding.ts (Lane G).
//
// All tests use memorySettingsBackend() — no Office.js dependency.
// refresh.ts is NOT tested here because it depends on writeResult (Lane F /
// Office.js).  Integration-style refresh tests belong in Lane J.

import { describe, it, expect, beforeEach } from "vitest";
import { memorySettingsBackend } from "../../src/connection/connectionStore.js";
import {
  type SavedQuery,
  newQueryId,
  saveQueryBinding,
  loadQueryBindings,
  deleteQueryBinding,
} from "../../src/excel/binding.js";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Build a minimal valid SavedQuery for testing. */
function makeQuery(overrides?: Partial<SavedQuery>): SavedQuery {
  return {
    queryId: newQueryId(),
    sql: "SELECT 1",
    rowCap: 1000,
    sheetName: "Sheet1",
    anchorAddress: "A1",
    endpointHost: "localhost",
    createdAt: new Date().toISOString(),
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// newQueryId
// ---------------------------------------------------------------------------

describe("newQueryId", () => {
  it("produces IDs starting with 'q_'", () => {
    const id = newQueryId();
    expect(id.startsWith("q_")).toBe(true);
  });

  it("produces unique IDs across many calls", () => {
    const ids = new Set(Array.from({ length: 500 }, () => newQueryId()));
    // All 500 should be distinct.
    expect(ids.size).toBe(500);
  });

  it("produces only string values", () => {
    expect(typeof newQueryId()).toBe("string");
  });
});

// ---------------------------------------------------------------------------
// loadQueryBindings — empty backend
// ---------------------------------------------------------------------------

describe("loadQueryBindings (empty backend)", () => {
  it("returns an empty array when nothing has been saved", () => {
    const backend = memorySettingsBackend();
    expect(loadQueryBindings(backend)).toEqual([]);
  });

  it("returns an empty array when the stored value is corrupt JSON", async () => {
    const backend = memorySettingsBackend();
    await backend.set("scx.queries", "NOT_VALID_JSON{{{{");
    expect(loadQueryBindings(backend)).toEqual([]);
  });

  it("returns an empty array when the stored value is a non-array JSON value", async () => {
    const backend = memorySettingsBackend();
    await backend.set("scx.queries", JSON.stringify({ not: "an array" }));
    expect(loadQueryBindings(backend)).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// saveQueryBinding / loadQueryBindings — round-trip
// ---------------------------------------------------------------------------

describe("saveQueryBinding / loadQueryBindings (round-trip)", () => {
  let backend: ReturnType<typeof memorySettingsBackend>;

  beforeEach(() => {
    backend = memorySettingsBackend();
  });

  it("persists a single query and reads it back", async () => {
    const q = makeQuery({ sql: "SELECT * FROM sales", rowCap: 5000 });
    await saveQueryBinding(q, backend);

    const loaded = loadQueryBindings(backend);
    expect(loaded).toHaveLength(1);
    expect(loaded[0]).toEqual(q);
  });

  it("appends multiple queries", async () => {
    const q1 = makeQuery({ sql: "SELECT 1" });
    const q2 = makeQuery({ sql: "SELECT 2" });
    await saveQueryBinding(q1, backend);
    await saveQueryBinding(q2, backend);

    const loaded = loadQueryBindings(backend);
    expect(loaded).toHaveLength(2);
    expect(loaded.map((r) => r.sql)).toEqual(["SELECT 1", "SELECT 2"]);
  });

  it("JSON round-trips all fields without loss", async () => {
    const q = makeQuery({
      sql: "SELECT id, ts FROM events WHERE ts > '2024-01-01'",
      rowCap: 10000,
      sheetName: "Events Data",
      anchorAddress: "C3",
      endpointHost: "spark.example.com",
      createdAt: "2024-06-01T12:00:00.000Z",
    });
    await saveQueryBinding(q, backend);

    const loaded = loadQueryBindings(backend);
    expect(loaded[0]).toStrictEqual(q);
  });

  it("preserves unicode in SQL and sheet names", async () => {
    const q = makeQuery({
      sql: "SELECT '日本語' AS text",
      sheetName: "販売データ",
    });
    await saveQueryBinding(q, backend);
    const loaded = loadQueryBindings(backend);
    expect(loaded[0]?.sql).toBe("SELECT '日本語' AS text");
    expect(loaded[0]?.sheetName).toBe("販売データ");
  });
});

// ---------------------------------------------------------------------------
// Upsert semantics
// ---------------------------------------------------------------------------

describe("saveQueryBinding upsert", () => {
  let backend: ReturnType<typeof memorySettingsBackend>;

  beforeEach(() => {
    backend = memorySettingsBackend();
  });

  it("replaces the existing record when queryId already exists", async () => {
    const id = newQueryId();
    const original = makeQuery({ queryId: id, sql: "SELECT 1" });
    const updated = makeQuery({ queryId: id, sql: "SELECT 2", rowCap: 9999 });

    await saveQueryBinding(original, backend);
    await saveQueryBinding(updated, backend);

    const loaded = loadQueryBindings(backend);
    // Still only one entry.
    expect(loaded).toHaveLength(1);
    expect(loaded[0]?.sql).toBe("SELECT 2");
    expect(loaded[0]?.rowCap).toBe(9999);
  });

  it("preserves the position of the updated record in the array", async () => {
    const idA = newQueryId();
    const idB = newQueryId();
    const idC = newQueryId();
    await saveQueryBinding(makeQuery({ queryId: idA, sql: "A" }), backend);
    await saveQueryBinding(makeQuery({ queryId: idB, sql: "B" }), backend);
    await saveQueryBinding(makeQuery({ queryId: idC, sql: "C" }), backend);

    // Update B in the middle.
    await saveQueryBinding(makeQuery({ queryId: idB, sql: "B-updated" }), backend);

    const loaded = loadQueryBindings(backend);
    expect(loaded).toHaveLength(3);
    expect(loaded[0]?.queryId).toBe(idA);
    expect(loaded[1]?.queryId).toBe(idB);
    expect(loaded[1]?.sql).toBe("B-updated");
    expect(loaded[2]?.queryId).toBe(idC);
  });

  it("does not duplicate when the same object is saved twice", async () => {
    const q = makeQuery();
    await saveQueryBinding(q, backend);
    await saveQueryBinding(q, backend);
    expect(loadQueryBindings(backend)).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// deleteQueryBinding
// ---------------------------------------------------------------------------

describe("deleteQueryBinding", () => {
  let backend: ReturnType<typeof memorySettingsBackend>;

  beforeEach(() => {
    backend = memorySettingsBackend();
  });

  it("removes the matching query", async () => {
    const q1 = makeQuery();
    const q2 = makeQuery();
    await saveQueryBinding(q1, backend);
    await saveQueryBinding(q2, backend);

    await deleteQueryBinding(q1.queryId, backend);

    const loaded = loadQueryBindings(backend);
    expect(loaded).toHaveLength(1);
    expect(loaded[0]?.queryId).toBe(q2.queryId);
  });

  it("is a no-op (no error) when the ID does not exist", async () => {
    const q = makeQuery();
    await saveQueryBinding(q, backend);

    await expect(
      deleteQueryBinding("non-existent-id", backend)
    ).resolves.toBeUndefined();

    // Original record still present.
    expect(loadQueryBindings(backend)).toHaveLength(1);
  });

  it("leaves an empty array after deleting the only record", async () => {
    const q = makeQuery();
    await saveQueryBinding(q, backend);
    await deleteQueryBinding(q.queryId, backend);
    expect(loadQueryBindings(backend)).toEqual([]);
  });

  it("is a no-op when the backend is empty", async () => {
    await expect(
      deleteQueryBinding("q_anything", backend)
    ).resolves.toBeUndefined();
    expect(loadQueryBindings(backend)).toEqual([]);
  });

  it("removes only the targeted record when multiple queries exist", async () => {
    const queries = [makeQuery(), makeQuery(), makeQuery()];
    for (const q of queries) {
      await saveQueryBinding(q, backend);
    }

    await deleteQueryBinding(queries[1]!.queryId, backend);

    const loaded = loadQueryBindings(backend);
    expect(loaded).toHaveLength(2);
    expect(loaded.map((r) => r.queryId)).toEqual([
      queries[0]!.queryId,
      queries[2]!.queryId,
    ]);
  });
});

// ---------------------------------------------------------------------------
// DECISIONS #6 guard: SavedQuery MUST NOT carry a token field
// ---------------------------------------------------------------------------

describe("DECISIONS #6 — no token on SavedQuery", () => {
  it("the SavedQuery interface has no 'token' property at runtime", async () => {
    const backend = memorySettingsBackend();
    const q = makeQuery();
    await saveQueryBinding(q, backend);

    const loaded = loadQueryBindings(backend);
    expect(loaded).toHaveLength(1);

    const record = loaded[0] as unknown as Record<string, unknown>;
    // Explicit check: "token" must not appear on a persisted record.
    expect(Object.prototype.hasOwnProperty.call(record, "token")).toBe(false);
    // "password" and "secret" variants too.
    expect(Object.prototype.hasOwnProperty.call(record, "password")).toBe(false);
    expect(Object.prototype.hasOwnProperty.call(record, "secret")).toBe(false);
    expect(Object.prototype.hasOwnProperty.call(record, "bearerToken")).toBe(false);
    expect(Object.prototype.hasOwnProperty.call(record, "authToken")).toBe(false);
  });

  it("SavedQuery only exposes the documented non-secret fields", async () => {
    const backend = memorySettingsBackend();
    const q = makeQuery();
    await saveQueryBinding(q, backend);

    const loaded = loadQueryBindings(backend);
    const record = loaded[0]!;
    const keys = Object.keys(record).sort();

    // Exactly these seven fields — no extras.
    expect(keys).toEqual(
      [
        "anchorAddress",
        "createdAt",
        "endpointHost",
        "queryId",
        "rowCap",
        "sheetName",
        "sql",
      ].sort()
    );
  });

  it("endpointHost stores only the host string, not a full URI or token", async () => {
    const backend = memorySettingsBackend();
    const q = makeQuery({ endpointHost: "spark.example.com" });
    await saveQueryBinding(q, backend);

    const loaded = loadQueryBindings(backend);
    const host = loaded[0]?.endpointHost ?? "";

    // Must not look like a URI (no scheme, no port with colon, no path separator
    // that would indicate a full connection string).
    expect(host).not.toMatch(/^sc:\/\//);
    expect(host).not.toMatch(/^https?:\/\//);
    expect(host).toBe("spark.example.com");
  });
});

// ---------------------------------------------------------------------------
// Corrupt / partial records are silently filtered
// ---------------------------------------------------------------------------

describe("corrupt stored records are filtered on load", () => {
  it("drops records that are missing required fields", async () => {
    const backend = memorySettingsBackend();
    // Manually write an array containing one valid and one corrupt entry.
    const valid = makeQuery();
    const corrupt = { queryId: "bad", sql: "SELECT 1" }; // missing most fields
    await backend.set(
      "scx.queries",
      JSON.stringify([valid, corrupt])
    );

    const loaded = loadQueryBindings(backend);
    expect(loaded).toHaveLength(1);
    expect(loaded[0]?.queryId).toBe(valid.queryId);
  });
});
