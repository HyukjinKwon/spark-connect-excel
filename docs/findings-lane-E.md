<!-- SPDX-License-Identifier: Apache-2.0 -->

# findings-lane-E.md - Query UX (Lane E)

## Files delivered

| File | Purpose |
|------|---------|
| `src/taskpane/taskpane.ts` | Entry point: `Office.onReady -> boot()`. Thin - only lifecycle + wiring. |
| `src/taskpane/queryPanel.ts` | All UI logic: connection, SQL editor, run/cancel, saved queries, chart. |
| `src/taskpane/ui.css` | Stylesheet for the narrow (~320 px) task pane. Spark orange `#E25A1C` accent. |

## Run flow

1. **Boot** (`taskpane.ts`):
   - `Office.onReady` fires -> `boot()`.
   - Renders the pane chrome (header + scrollable body).
   - Calls `createDialogBridge(dialog.html)` to open the COI host dialog window.
   - Shows an `sc-engine-banner` while `bridge.ensureReady()` runs in the background.
   - `renderQueryPanel(body, bridge)` is called immediately so the user can configure the connection while Pyodide boots.

2. **Connection** (`queryPanel.ts -> renderConnectionForm`):
   - Lane I's `renderConnectionForm` owns all form controls.
   - On submit: `saveConnection` (non-secret config) + `saveToken` (secure); then `bridge.ensureReady()` + `bridge.connect(uri, {token})`.
   - On success: the Connection section auto-collapses.

3. **Run flow** (`queryPanel.ts -> handleRun`):
   - Validates SQL + row cap locally.
   - `bridge.ensureReady()` (idempotent) -> `bridge.connect(...)` -> `bridge.runSQL(sql, rowCap)`.
   - `writeResult(result, {})` -> `WriteResultInfo`.
   - `saveQueryBinding({ queryId, sql, rowCap, sheetName, anchorAddress, endpointHost, createdAt })` - token intentionally absent (DECISIONS #6).
   - Success: row count + truncation notice; "Insert chart" button appears.

4. **Insert chart**: `insertChart(lastWriteInfo, lastResultSchema)` on button click.

5. **Refresh single**: `refreshQuery(queryId, bridge)` -> `RefreshResult`.

6. **Refresh All**: `refreshAll(bridge)` -> `RefreshResult[]`; each item status updated in-place.

## Seam assumptions

- **Lane D `SparkBridgeClient`**: `createDialogBridge(url)` opens the dialog. `BridgeEvent { event: "progress", payload: string }` updates the engine banner; `event: "status" | "ready"` triggers a status badge refresh.
- **Lane F `writeResult(result, {})`**: empty options object means "write at active cell". Returns `WriteResultInfo { sheetName: string; headerRangeAddress: string }`.
- **Lane H `insertChart(info, schema, opts?)`**: called with `WriteResultInfo` + `ColumnMeta[]`.
- **Lane G `SavedQuery`**: `{ queryId, sql, rowCap, sheetName, anchorAddress, endpointHost, createdAt }`.
- **Lane G `refreshQuery` / `refreshAll`**: return `RefreshResult` / `RefreshResult[]` with `{ queryId, ok, rowCount, truncated, error? }`.
- **Lane I `renderConnectionForm`**: `onSubmit(cfg: ConnectionConfig, token?: string) => void`.

## DECISIONS compliance

| Decision | How Lane E satisfies it |
|----------|------------------------|
| #4 SQL surface only | Only `runSQL` is exposed; no DataFrame/Python UI. |
| #5 Row cap default 10k | `DEFAULT_ROW_CAP = 10_000`; overridable per-query. |
| #6 Token never in a cell | Token goes through `saveToken`/`loadToken` only; `saveQueryBinding` stores `endpointHost`, not the token. |
| #7 No runtime imports | `src/taskpane/**` imports nothing from `src/runtime/**`. All Spark I/O via `SparkBridgeClient`. |

## Notable design choices

- **Background `ensureReady()`**: Pyodide boot is kicked off on page load in the background; the engine banner shows progress. The UI is not blocked.
- **Dialog opened eagerly**: the dialog is opened once at boot (not lazily on first Run) to hide Pyodide's cold-start latency.
- **No global state**: the panel API is stored in a `WeakMap` keyed on the body element rather than a module-level variable, keeping the module testable.
- **Framework-free DOM**: `el(tag, attrs, text?)` + `div(className, ...children)` - no React/Vue/Lit.
