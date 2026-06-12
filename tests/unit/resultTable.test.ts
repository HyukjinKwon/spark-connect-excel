// SPDX-License-Identifier: Apache-2.0
import { describe, it, expect } from "vitest";
import { renderResultTable } from "../../src/demo/resultTable";
import type { SparkResult } from "../../src/seam";

const base: SparkResult = {
  schema: [
    { name: "id", type: "bigint" },
    { name: "name", type: "string" },
  ],
  rows: [
    [1, "a"],
    [2, null],
  ],
  rowCount: 2,
  truncated: false,
};

describe("renderResultTable", () => {
  it("renders a header cell per column with name and type", () => {
    const node = renderResultTable(base);
    const ths = node.querySelectorAll("thead th");
    expect(ths).toHaveLength(2);
    expect(ths[0].querySelector(".demo-table__name")?.textContent).toBe("id");
    expect(ths[0].querySelector(".demo-table__type")?.textContent).toBe("bigint");
  });

  it("renders one body row per data row", () => {
    const node = renderResultTable(base);
    expect(node.querySelectorAll("tbody tr")).toHaveLength(2);
  });

  it("renders null cells as NULL with the null class", () => {
    const node = renderResultTable(base);
    const nullCell = node.querySelector(".demo-table__null");
    expect(nullCell?.textContent).toBe("NULL");
  });

  it("right-aligns numeric cells", () => {
    const node = renderResultTable(base);
    const firstCell = node.querySelector("tbody tr td");
    expect(firstCell?.classList.contains("demo-table__num")).toBe(true);
  });

  it("shows a meta line with the row and column counts", () => {
    const node = renderResultTable(base);
    expect(node.querySelector(".demo-result__meta")?.textContent).toBe("2 rows × 2 columns");
  });

  it("shows a truncation banner only when truncated", () => {
    expect(renderResultTable(base).querySelector(".demo-result__banner")).toBeNull();
    const truncated = renderResultTable({ ...base, truncated: true });
    expect(truncated.querySelector(".demo-result__banner")?.textContent).toContain("truncated");
  });

  it("singularizes the meta line for a single row/column", () => {
    const one: SparkResult = {
      schema: [{ name: "x", type: "int" }],
      rows: [[1]],
      rowCount: 1,
      truncated: false,
    };
    expect(renderResultTable(one).querySelector(".demo-result__meta")?.textContent).toBe(
      "1 row × 1 column",
    );
  });
});
