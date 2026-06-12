<!-- SPDX-License-Identifier: Apache-2.0 -->

# Lane G findings - Refresh & Binding

## Settings key and schema

All saved query bindings are stored under a **single document-settings key**:

```
"scx.queries"
```

The value is a JSON array of `SavedQuery` objects.  Using one key (rather than
one key per query) keeps the settings namespace flat and makes atomic reads
cheap - Lane F/E/H can call `loadQueryBindings()` once and iterate.

### SavedQuery shape

```ts
interface SavedQuery {
  queryId:       string;   // "q_" + base36-timestamp + "_" + 6 hex digits
  sql:           string;   // SQL text to re-execute
  rowCap:        number;   // row cap forwarded to bridge.runSQL
  sheetName:     string;   // destination sheet
  anchorAddress: string;   // top-left cell of the result range, e.g. "A1"
  endpointHost:  string;   // bare host ONLY - no scheme, no port, no token
  createdAt:     string;   // ISO-8601 UTC timestamp
}
```

DECISIONS #6 invariant (tested): SavedQuery carries exactly these seven fields.
token, password, secret, bearerToken, and authToken are all absent.
endpointHost is the host component only (e.g. "spark.example.com") - NOT a
full sc:// URI, NOT a connection string.  The bearer token lives in
OfficeRuntime.storage (Lane I) and is fetched at connection time.

## Persistence operations

| Function | Semantics |
|----------|-----------|
| saveQueryBinding(q, backend?) | Upsert by queryId: replaces in-place if the ID already exists (position preserved), appends otherwise. |
| loadQueryBindings(backend?) | Synchronous read from backend in-memory cache.  Returns [] on fresh workbook or corrupt value (fail-safe). |
| deleteQueryBinding(queryId, backend?) | Removes the matching record.  No-op (no throw) if the ID is not found. |
| newQueryId() | Generates a collision-resistant ID: "q_" + Date.now().toString(36) + "_" + random 6-hex. |

All functions accept an optional SettingsBackend (imported from
src/connection/connectionStore.ts).  The default is officeDocumentSettingsBackend().
Unit tests inject memorySettingsBackend() - no Office.js required.

## Refresh semantics

### refreshQuery(queryId, bridge, backend?)

1. loadQueryBindings(backend) - find the record by queryId.
2. Throw a descriptive Error if not found (lists available IDs).
3. await bridge.runSQL(q.sql, q.rowCap) - SparkResult.
4. await writeResult(result, { anchorAddress, sheetName }) (Lane F) - rewrites the SAME range.
5. Returns WriteResultInfo from Lane F.

### refreshAll(bridge, backend?)

- Loads all bindings, runs refreshQuery on each via Promise.allSettled.
- One failure does NOT abort the rest (best-effort, Refresh All UX).
- Returns { queryId, ok, error? }[] - one entry per saved query, in load order.

## Dependency notes

- SettingsBackend, officeDocumentSettingsBackend(), memorySettingsBackend()
  are imported (not duplicated) from src/connection/connectionStore.ts (Lane I).
- writeResult and WriteResultInfo are imported from src/excel/rangeWriter.ts
  (Lane F).  Until Lane F delivers that file refresh.ts will not typecheck -
  expected in parallel development, resolves at integration.
- SparkBridge / SparkResult imported from src/seam.ts (frozen integrator seam).

## Files created

| File | Purpose |
|------|---------|
| src/excel/binding.ts | SavedQuery, persistence functions, newQueryId |
| src/excel/refresh.ts | refreshQuery, refreshAll |
| tests/unit/binding.test.ts | Vitest unit tests (pure, no Office.js) |
| docs/findings-lane-G.md | This file |
