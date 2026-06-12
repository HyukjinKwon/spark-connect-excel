<!-- SPDX-License-Identifier: Apache-2.0 -->

# API_CONTRACT.md — the frozen seam

The stable interface between **task-pane lanes** (E/F/G/H) and **runtime lanes**
(B/C/D). Canonical source of the types is [`src/seam.ts`](src/seam.ts); this file
explains them. Do not change a signature without announcing it in
`COORDINATION.md` first.

## 1. The bridge (`src/seam.ts` → `SparkBridge`)

```ts
ensureReady(): Promise<void>            // Pyodide + pcw.install() + COI verified
connect(uri, { token? }): Promise<void> // build/refresh the SparkSession
runSQL(sql, rowCap): Promise<SparkResult>
schemaOf(sql): Promise<ColumnMeta[]>    // schema only, no data
status(): RuntimeStatus                  // cheap, synchronous
cancel(): void
```

Implemented twice against the same interface:

- **`SparkBridgeHost`** (dialog side, Lane D) — does the real work by calling the
  Python runtime through `window.__pcwRunPython(src)` (Lane C).
- **`SparkBridgeClient`** (task-pane side, Lane D) — forwards each call across the
  Office dialog channel using the envelope below, and resolves on the reply.

## 2. The message envelope (task pane ⟷ dialog)

Office dialog messages are **strings**. Every message is a JSON `BridgeMessage`
(`req` | `res` | `evt`). The client sends `BridgeRequest`, the host replies with a
matching `BridgeResponse` (by `id`); the host may also push `BridgeEvent`s
(boot progress, status). See `encodeMessage` / `decodeMessage` in `src/seam.ts`.

Transport: parent→dialog `dialog.messageChild(str)`; dialog→parent
`Office.context.ui.messageParent(str)`; parent listens via
`Office.EventType.DialogMessageReceived`.

## 3. The Python runtime contract (`python/spark_excel_runtime.py`, Lane D)

A pure-Python module loaded into Pyodide. It exposes module-level functions the
dialog drives via tiny `__pcwRunPython` snippets. Every function **returns a JSON
string** (DECISIONS #8):

```python
connect(uri: str, token: str | None) -> str   # '{"ok": true}' or '{"ok": false, "error": {...}}'
run_sql(sql: str, row_cap: int)       -> str   # JSON shaped exactly like SparkResult
schema_of(sql: str)                   -> str   # JSON: {"schema": [{"name","type"}, ...]}
```

`run_sql` does, in effect:

```python
df = spark.sql(sql).limit(row_cap + 1)
pdf = df.toPandas()
truncated = len(pdf) > row_cap
pdf = pdf.iloc[:row_cap]
# -> {schema: [{name,type}...], rows: [[...]], rowCount, truncated}
```

Spark types come from `df.schema` (authoritative); values are JSON-normalised
(timestamps/dates → ISO-8601 strings, NaN/NaT → null). This module is unit-
tested with pytest using a fake `spark` (no live cluster) — Lane D owns the fake.

## 4. Excel-side contracts (task-pane lanes)

- **Range writer (F):** `writeResult(result: SparkResult, anchor: Excel.Range)` —
  header row + typed cells; sets number formats from `ColumnMeta.type`; renders a
  truncation banner above the anchor when `result.truncated`.
- **Binding (G):** persists `{ queryId, sql, rowCap, sheet, anchorAddress, endpointHost }`
  to `Office document settings`; never the token (DECISIONS #6). `refresh(queryId)`
  re-runs `runSQL` and rewrites the same range.
- **Chart (H):** `insertChart(result, range)` picks a chart type from the schema
  (one numeric + one categorical → column; time + numeric → line; etc.) and calls
  `worksheet.charts.add`.
- **Connection (I):** owns the `uri`/`token` capture + `RuntimeStatus` surfacing;
  hands `connect()` its arguments. Token storage per DECISIONS #6.
