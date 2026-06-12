// SPDX-License-Identifier: Apache-2.0
//
// queryPanel.ts — Query UX panel (Lane E).
//
// Composes connection form, SQL editor, row-cap, run/cancel, saved-query list,
// and chart insertion into a single DOM tree rendered into the provided root.
//
// Import policy (DECISIONS #7): MUST NOT import from src/runtime/**.
// All Spark operations go through SparkBridgeClient.

import type { SparkBridgeClient } from "../bridge/sparkBridgeClient.js";
import {
  renderConnectionForm,
  connectionStatusBadge,
  validateConnection,
} from "../connection/connectionForm.js";
import type { ConnectionConfig } from "../connection/connectionStore.js";
import {
  buildRemoteUri,
  saveConnection,
  loadConnection,
  saveToken,
  loadToken,
  officeDocumentSettingsBackend,
} from "../connection/connectionStore.js";
import type { WriteResultInfo } from "../excel/rangeWriter.js";
import { writeResult } from "../excel/rangeWriter.js";
import { insertChart } from "../excel/chart.js";
import type { SavedQuery } from "../excel/binding.js";
import { saveQueryBinding, loadQueryBindings, newQueryId } from "../excel/binding.js";
import { refreshQuery, refreshAll } from "../excel/refresh.js";
import type { ColumnMeta, RuntimeStatus } from "../seam.js";

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const DEFAULT_ROW_CAP = 10_000;

// ---------------------------------------------------------------------------
// Tiny framework-free DOM helper
// ---------------------------------------------------------------------------

type ElAttrs = Record<string, string | boolean | number>;

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: ElAttrs = {},
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "htmlFor") {
      (node as HTMLLabelElement).htmlFor = String(v);
    } else if (k === "className") {
      node.className = String(v);
    } else if (typeof v === "boolean") {
      if (v) node.setAttribute(k, "");
    } else {
      node.setAttribute(k, String(v));
    }
  }
  if (text !== undefined) node.textContent = text;
  return node;
}

function div(className: string, ...children: (HTMLElement | undefined)[]): HTMLDivElement {
  const d = el("div", { className });
  for (const c of children) {
    if (c) d.appendChild(c);
  }
  return d;
}

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

type RunState = "idle" | "starting" | "running" | "done" | "error";

// ---------------------------------------------------------------------------
// Panel API — returned to taskpane.ts for event forwarding.
// ---------------------------------------------------------------------------

export interface QueryPanelApi {
  /** Re-render the status badge row from the current bridge status. */
  refreshStatus(): void;
}

// ---------------------------------------------------------------------------
// renderQueryPanel
// ---------------------------------------------------------------------------

/**
 * Render the full query UX into `root`, wiring it to `bridge`.
 *
 * Returns a `{ api, dispose }` pair.  `api.refreshStatus()` should be called
 * whenever a BridgeEvent arrives from the dialog.  `dispose()` removes all
 * event listeners attached by this function.
 */
export function renderQueryPanel(
  root: HTMLElement,
  bridge: SparkBridgeClient,
): { api: QueryPanelApi; dispose: () => void } {
  while (root.firstChild) root.removeChild(root.firstChild);

  const backend = officeDocumentSettingsBackend();
  const disposers: Array<() => void> = [];

  // Load persisted connection and token.
  let activeCfg: ConnectionConfig | null = loadConnection(backend);
  let activeToken: string | undefined;
  void loadToken().then((t) => {
    if (t) activeToken = t;
  });

  // ── Status bar (always visible, above sections) ───────────────────────────

  const statusBar = div("sc-status-bar");
  renderStatusBadges(statusBar, bridge.status());

  function refreshStatus(): void {
    renderStatusBadges(statusBar, bridge.status());
  }

  // ── Connection section ────────────────────────────────────────────────────

  const connBody = div("sc-section__body");
  const connSection = buildCollapsibleSection("Connection", connBody, !activeCfg);

  // Sub-status feedback for connect-in-progress.
  const connStatusEl = el("p", { className: "sc-info sc-mt-4" });
  connStatusEl.style.display = "none";

  // Lane I's renderConnectionForm handles all input controls.
  const disposeForm = renderConnectionForm(connBody, activeCfg, (cfg, token) => {
    void (async () => {
      activeCfg = cfg;
      if (token !== undefined) {
        activeToken = token;
        await saveToken(token).catch(() => undefined);
      }
      await saveConnection(cfg, backend).catch(() => undefined);

      connStatusEl.style.color = "";
      connStatusEl.textContent = "Connecting…";
      connStatusEl.style.display = "";

      try {
        await bridge.ensureReady();
        await bridge.connect(buildRemoteUri(cfg), { token: activeToken });
        connStatusEl.textContent = "Connected.";
        refreshStatus();
        collapseSection(connSection);
      } catch (err) {
        connStatusEl.textContent = `Connection failed: ${errMsg(err)}`;
        connStatusEl.style.color = "#721c24";
        refreshStatus();
      }
    })();
  });
  disposers.push(disposeForm);
  connBody.appendChild(connStatusEl);

  // ── Query section ─────────────────────────────────────────────────────────

  const queryBody = div("sc-section__body");

  // SQL textarea.
  const sqlLabel = el("label", { className: "sc-label", htmlFor: "sc-sql" }, "SQL");
  const sqlArea = el("textarea", {
    id: "sc-sql",
    className: "sc-textarea",
    placeholder: "SELECT * FROM my_table LIMIT 100",
    rows: "5",
    spellcheck: "false",
    autocomplete: "off",
  });

  // Row cap.
  const capLabel = el(
    "label",
    {
      className: "sc-label",
      htmlFor: "sc-rowcap",
    },
    "Row cap",
  );
  const capInput = el("input", {
    id: "sc-rowcap",
    type: "number",
    className: "sc-input",
    min: "1",
    max: "1000000",
    value: String(DEFAULT_ROW_CAP),
  });

  // Destination hint — shows the active cell address.
  const destHint = el("p", { className: "sc-dest-hint" });
  setDestHint(destHint);

  // Error and success message areas.
  const queryErrorEl = el("div", {
    className: "sc-error",
    role: "alert",
    "aria-live": "assertive",
  });
  const querySuccessEl = el("div", {
    className: "sc-success",
    role: "status",
    "aria-live": "polite",
  });

  // "Insert chart" button — only visible after a successful run.
  const chartBtn = el(
    "button",
    {
      className: "sc-btn sc-btn--secondary sc-hidden",
      type: "button",
    },
    "Insert chart",
  );
  let lastWriteInfo: WriteResultInfo | null = null;
  let lastResultSchema: ColumnMeta[] | null = null;

  function onChartClick(): void {
    if (!lastWriteInfo || !lastResultSchema) return;
    chartBtn.disabled = true;
    void insertChart(lastWriteInfo, lastResultSchema)
      .then(() => {
        setText(querySuccessEl, "Chart inserted.");
      })
      .catch((err: unknown) => {
        setText(queryErrorEl, errMsg(err));
      })
      .finally(() => {
        chartBtn.disabled = false;
      });
  }
  chartBtn.addEventListener("click", onChartClick);
  disposers.push(() => chartBtn.removeEventListener("click", onChartClick));

  // Run / Cancel buttons.
  const runBtn = el(
    "button",
    {
      className: "sc-btn sc-btn--primary",
      type: "button",
    },
    "Run",
  );
  const cancelBtn = el(
    "button",
    {
      className: "sc-btn sc-btn--secondary",
      type: "button",
      disabled: true,
    },
    "Cancel",
  );

  // runState is mutated by setRunState; the unused-var linter must not
  // flag it, so we reference it via the void pattern inside setRunState.
  let runState: RunState = "idle";

  function setRunState(state: RunState): void {
    runState = state;
    void runState; // suppress noUnusedLocals
    const busy = state === "starting" || state === "running";
    (runBtn as HTMLButtonElement).disabled = busy;
    (cancelBtn as HTMLButtonElement).disabled = !busy;

    if (state === "starting") {
      runBtn.innerHTML = '<span class="sc-spinner"></span> Starting…';
    } else if (state === "running") {
      runBtn.innerHTML = '<span class="sc-spinner"></span> Running…';
    } else {
      runBtn.textContent = "Run";
    }
  }

  async function handleRun(): Promise<void> {
    const sql = sqlArea.value.trim();
    if (!sql) {
      setText(queryErrorEl, "Please enter a SQL query.");
      return;
    }
    if (!activeCfg) {
      setText(queryErrorEl, "Please configure a connection first (expand the Connection section).");
      return;
    }
    const connErr = validateConnection(activeCfg.host, activeCfg.port);
    if (connErr) {
      setText(queryErrorEl, connErr);
      return;
    }
    const rowCap = parseInt(capInput.value, 10);
    if (!Number.isFinite(rowCap) || rowCap < 1) {
      setText(queryErrorEl, "Row cap must be a positive integer.");
      return;
    }

    // Reset UI state.
    setText(queryErrorEl, "");
    setText(querySuccessEl, "");
    chartBtn.classList.add("sc-hidden");
    lastWriteInfo = null;
    lastResultSchema = null;
    setDestHint(destHint);
    setRunState("starting");

    try {
      // 1. Boot Pyodide + pcw (idempotent — resolves immediately if already ready).
      await bridge.ensureReady();
      refreshStatus();

      // 2. Connect / re-connect.
      setRunState("running");
      await bridge.connect(buildRemoteUri(activeCfg), { token: activeToken });
      refreshStatus();

      // 3. Run SQL.
      const result = await bridge.runSQL(sql, rowCap);
      refreshStatus();

      // 4. Write result to the active cell range (Lane F).
      const info = await writeResult(result, {});
      lastWriteInfo = info;
      lastResultSchema = result.schema;

      // 5. Persist query binding (DECISIONS #6 — token NOT stored here).
      await saveQueryBinding({
        queryId: newQueryId(),
        sql,
        rowCap,
        sheetName: info.sheetName,
        anchorAddress: info.headerRangeAddress,
        endpointHost: activeCfg.host,
        createdAt: new Date().toISOString(),
      });

      // 6. Surface result feedback.
      const truncMsg = result.truncated ? ` (truncated at cap ${rowCap.toLocaleString()})` : "";
      setText(
        querySuccessEl,
        `${result.rowCount.toLocaleString()} rows written to ${info.sheetName}${truncMsg}.`,
      );
      chartBtn.classList.remove("sc-hidden");
      setRunState("done");

      // 7. Reload saved-queries list.
      await reloadSavedQueries();
    } catch (err) {
      setRunState("error");
      setText(queryErrorEl, errMsg(err));
      refreshStatus();
    }
  }

  function onRunClick(): void {
    void handleRun();
  }
  function onCancelClick(): void {
    bridge.cancel();
    setRunState("idle");
    setText(querySuccessEl, "");
    setText(queryErrorEl, "Cancelled.");
  }

  runBtn.addEventListener("click", onRunClick);
  cancelBtn.addEventListener("click", onCancelClick);
  disposers.push(() => runBtn.removeEventListener("click", onRunClick));
  disposers.push(() => cancelBtn.removeEventListener("click", onCancelClick));

  // Assemble query section body.
  queryBody.appendChild(div("sc-form-row", sqlLabel, div("sc-editor-wrap", sqlArea)));
  queryBody.appendChild(div("sc-cap-row", capLabel, capInput));
  queryBody.appendChild(destHint);
  queryBody.appendChild(queryErrorEl);
  queryBody.appendChild(querySuccessEl);
  queryBody.appendChild(div("sc-btn-row sc-mt-4", chartBtn));
  queryBody.appendChild(div("sc-btn-row sc-mt-4", runBtn, cancelBtn));

  // ── Saved queries section ─────────────────────────────────────────────────

  const savedQueriesBody = div("sc-section__body");

  async function reloadSavedQueries(): Promise<void> {
    await renderSavedQueriesContent(savedQueriesBody, bridge);
  }

  void reloadSavedQueries();

  // ── Assemble root ─────────────────────────────────────────────────────────

  const querySection = buildCollapsibleSection("SQL Query", queryBody, true);
  const savedSection = buildCollapsibleSection("Saved Queries", savedQueriesBody, true);

  root.appendChild(statusBar);
  root.appendChild(connSection);
  root.appendChild(querySection);
  root.appendChild(savedSection);

  // Auto-collapse connection section when we already have a persisted config.
  if (activeCfg) {
    collapseSection(connSection);
  }

  return {
    api: { refreshStatus },
    dispose: () => {
      for (const d of disposers) d();
    },
  };
}

// ---------------------------------------------------------------------------
// Saved queries content
// ---------------------------------------------------------------------------

async function renderSavedQueriesContent(
  container: HTMLElement,
  bridge: SparkBridgeClient,
): Promise<void> {
  while (container.firstChild) container.removeChild(container.firstChild);

  // loadQueryBindings is synchronous (cache-backed Office settings).
  let bindings: SavedQuery[];
  try {
    bindings = loadQueryBindings();
  } catch (err) {
    container.appendChild(
      el("p", { className: "sc-error" }, `Failed to load saved queries: ${errMsg(err)}`),
    );
    return;
  }

  if (bindings.length === 0) {
    container.appendChild(
      el("p", { className: "sc-query-empty" }, "No saved queries yet. Run a query to create one."),
    );
    return;
  }

  // Refresh-All button row.
  const refreshAllBtn = el(
    "button",
    { className: "sc-btn sc-btn--ghost sc-btn--sm", type: "button" },
    "Refresh All",
  );
  container.appendChild(div("sc-refresh-all-row", refreshAllBtn));

  // Query list.
  const listEl = div("sc-query-list");
  container.appendChild(listEl);

  // Map from queryId → status element, for Refresh-All result annotation.
  const itemStatusMap = new Map<string, HTMLElement>();

  for (const q of bindings) {
    const { itemEl, statusEl } = buildQueryItemEl(q, bridge, async () => {
      await renderSavedQueriesContent(container, bridge);
    });
    listEl.appendChild(itemEl);
    itemStatusMap.set(q.queryId, statusEl);
  }

  // Refresh-All handler.
  function onRefreshAll(): void {
    void (async () => {
      (refreshAllBtn as HTMLButtonElement).disabled = true;
      refreshAllBtn.innerHTML = '<span class="sc-spinner"></span> Refreshing…';

      try {
        // refreshAll returns { queryId, ok, error? }[] — no rowCount/truncated.
        const results = await refreshAll(bridge);
        for (const r of results) {
          const s = itemStatusMap.get(r.queryId);
          if (!s) continue;
          if (r.ok) {
            s.textContent = "Refreshed.";
            s.style.color = "#155724";
          } else {
            s.textContent = `Error: ${r.error ?? "unknown"}`;
            s.style.color = "#721c24";
          }
        }
      } catch (err) {
        container.insertBefore(
          el("p", { className: "sc-error" }, `Refresh-All failed: ${errMsg(err)}`),
          listEl,
        );
      } finally {
        (refreshAllBtn as HTMLButtonElement).disabled = false;
        refreshAllBtn.textContent = "Refresh All";
      }
    })();
  }

  refreshAllBtn.addEventListener("click", onRefreshAll);
}

function buildQueryItemEl(
  q: SavedQuery,
  bridge: SparkBridgeClient,
  onReload: () => Promise<void>,
): { itemEl: HTMLDivElement; statusEl: HTMLElement } {
  const itemEl = div("sc-query-item");

  // SQL preview (single line, ellipsis on overflow).
  const sqlEl = el("div", { className: "sc-query-item__sql", title: q.sql });
  sqlEl.textContent = q.sql;

  // Metadata line: date + host + anchor.
  const d = new Date(q.createdAt);
  const dateStr = Number.isNaN(d.getTime())
    ? q.createdAt
    : d.toLocaleString(undefined, { dateStyle: "short", timeStyle: "short" });
  const metaEl = el("div", { className: "sc-query-item__meta" });
  metaEl.textContent = `${dateStr} · ${q.endpointHost} · ${q.anchorAddress}`;

  const refreshBtn = el(
    "button",
    { className: "sc-btn sc-btn--ghost sc-btn--sm", type: "button" },
    "Refresh",
  );

  const statusEl = el("div", { className: "sc-query-item__status" });

  itemEl.appendChild(sqlEl);
  itemEl.appendChild(metaEl);
  itemEl.appendChild(div("sc-query-item__actions", refreshBtn));
  itemEl.appendChild(statusEl);

  function onRefresh(): void {
    void (async () => {
      (refreshBtn as HTMLButtonElement).disabled = true;
      refreshBtn.innerHTML = '<span class="sc-spinner"></span>';
      statusEl.textContent = "Refreshing…";
      statusEl.style.color = "";

      try {
        // refreshQuery returns WriteResultInfo on success.
        const info = await refreshQuery(q.queryId, bridge);
        statusEl.textContent =
          `Refreshed — ${info.rowCount.toLocaleString()} rows.` +
          (info.truncated ? " (truncated)" : "");
        statusEl.style.color = "#155724";
      } catch (err) {
        statusEl.textContent = `Error: ${errMsg(err)}`;
        statusEl.style.color = "#721c24";
      } finally {
        (refreshBtn as HTMLButtonElement).disabled = false;
        refreshBtn.textContent = "Refresh";
      }

      // Reload the list so any ordering/state changes are reflected.
      await onReload();
    })();
  }

  refreshBtn.addEventListener("click", onRefresh);
  return { itemEl, statusEl };
}

// ---------------------------------------------------------------------------
// Collapsible section builder
// ---------------------------------------------------------------------------

function buildCollapsibleSection(
  title: string,
  body: HTMLElement,
  expanded: boolean,
): HTMLDivElement {
  const section = el("div", { className: "sc-section" });

  const head = el("div", { className: "sc-section__head" });
  head.setAttribute("role", "button");
  head.setAttribute("tabindex", "0");
  head.setAttribute("aria-expanded", expanded ? "true" : "false");

  const titleEl = el("span", { className: "sc-section__title" }, title);
  const chevron = el("span", { className: "sc-section__chevron" }, "▼");
  head.appendChild(titleEl);
  head.appendChild(chevron);

  if (!expanded) body.classList.add("sc-collapsed");

  function toggle(): void {
    const open = head.getAttribute("aria-expanded") === "true";
    head.setAttribute("aria-expanded", open ? "false" : "true");
    body.classList.toggle("sc-collapsed", open);
  }

  head.addEventListener("click", toggle);
  head.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      toggle();
    }
  });

  section.appendChild(head);
  section.appendChild(body);
  return section;
}

function collapseSection(section: HTMLElement): void {
  const head = section.querySelector<HTMLElement>(".sc-section__head");
  const body = section.querySelector<HTMLElement>(".sc-section__body");
  if (head) head.setAttribute("aria-expanded", "false");
  if (body) body.classList.add("sc-collapsed");
}

// ---------------------------------------------------------------------------
// Status badge helper (delegates to Lane I's connectionStatusBadge)
// ---------------------------------------------------------------------------

function renderStatusBadges(root: HTMLElement, status: RuntimeStatus): void {
  connectionStatusBadge(root, status);
}

// ---------------------------------------------------------------------------
// Destination hint
// ---------------------------------------------------------------------------

function setDestHint(hintEl: HTMLElement): void {
  if (typeof Excel === "undefined") {
    hintEl.textContent = "Output will start at the active cell.";
    return;
  }
  Excel.run(async (ctx) => {
    const cell = ctx.workbook.getActiveCell();
    cell.load("address");
    await ctx.sync();
    hintEl.textContent = `Output will start at ${cell.address}.`;
  }).catch(() => {
    hintEl.textContent = "Output will start at the active cell.";
  });
}

// ---------------------------------------------------------------------------
// Text / error helpers
// ---------------------------------------------------------------------------

function setText(node: HTMLElement, msg: string): void {
  node.textContent = msg;
}

function errMsg(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
