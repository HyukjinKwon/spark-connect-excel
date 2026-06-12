// SPDX-License-Identifier: Apache-2.0
//
// resultTable.ts — pure DOM renderer for a SparkResult (used by the standalone
// web demo, which has no Excel grid to write into). Kept free of any runtime /
// Office dependency so it is unit-testable under jsdom.

import type { SparkResult } from "../seam";

/** Render a SparkResult as a `<div>` containing an optional banner + a table. */
export function renderResultTable(result: SparkResult): HTMLElement {
  const wrap = document.createElement("div");
  wrap.className = "demo-result";

  if (result.truncated) {
    const banner = document.createElement("div");
    banner.className = "demo-result__banner";
    banner.textContent = `Showing first ${result.rowCount} rows (result truncated).`;
    wrap.appendChild(banner);
  }

  const meta = document.createElement("div");
  meta.className = "demo-result__meta";
  const cols = result.schema.length;
  meta.textContent =
    `${result.rowCount} row${result.rowCount === 1 ? "" : "s"} × ` +
    `${cols} column${cols === 1 ? "" : "s"}`;
  wrap.appendChild(meta);

  const table = document.createElement("table");
  table.className = "demo-table";

  // Header: column name + type.
  const thead = document.createElement("thead");
  const htr = document.createElement("tr");
  for (const col of result.schema) {
    const th = document.createElement("th");
    const name = document.createElement("span");
    name.className = "demo-table__name";
    name.textContent = col.name;
    const type = document.createElement("span");
    type.className = "demo-table__type";
    type.textContent = col.type;
    th.appendChild(name);
    th.appendChild(type);
    htr.appendChild(th);
  }
  thead.appendChild(htr);
  table.appendChild(thead);

  // Body.
  const tbody = document.createElement("tbody");
  for (const row of result.rows) {
    const tr = document.createElement("tr");
    for (const cell of row) {
      const td = document.createElement("td");
      if (cell === null || cell === undefined) {
        td.className = "demo-table__null";
        td.textContent = "NULL";
      } else {
        td.textContent = String(cell);
        if (typeof cell === "number") td.classList.add("demo-table__num");
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  wrap.appendChild(table);

  return wrap;
}
