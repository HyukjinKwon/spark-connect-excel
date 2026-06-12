// SPDX-License-Identifier: Apache-2.0
import { describe, it, expect } from "vitest";
import {
  renderChart,
  isNumeric,
  isTemporal,
  isCategorical,
} from "../../src/demo/chartView";
import type { SparkResult } from "../../src/seam";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeResult(
  schema: { name: string; type: string }[],
  rows: unknown[][],
): SparkResult {
  return { schema, rows, rowCount: rows.length, truncated: false };
}

// ---------------------------------------------------------------------------
// Classification helpers
// ---------------------------------------------------------------------------

describe("isNumeric", () => {
  it("returns true for exact numeric type names", () => {
    for (const t of ["bigint", "int", "integer", "smallint", "tinyint", "double", "float", "real"]) {
      expect(isNumeric(t)).toBe(true);
    }
  });

  it("returns true for decimal/numeric prefix types", () => {
    expect(isNumeric("decimal(10,2)")).toBe(true);
    expect(isNumeric("numeric(8,4)")).toBe(true);
  });

  it("is case-insensitive", () => {
    expect(isNumeric("BIGINT")).toBe(true);
    expect(isNumeric("Double")).toBe(true);
  });

  it("returns false for non-numeric types", () => {
    expect(isNumeric("string")).toBe(false);
    expect(isNumeric("timestamp")).toBe(false);
    expect(isNumeric("boolean")).toBe(false);
  });
});

describe("isTemporal", () => {
  it("returns true for date/timestamp types", () => {
    expect(isTemporal("date")).toBe(true);
    expect(isTemporal("timestamp")).toBe(true);
    expect(isTemporal("timestamp_ntz")).toBe(true);
    expect(isTemporal("timestamp_ltz")).toBe(true);
  });

  it("is case-insensitive", () => {
    expect(isTemporal("DATE")).toBe(true);
    expect(isTemporal("TIMESTAMP")).toBe(true);
  });

  it("returns false for non-temporal types", () => {
    expect(isTemporal("string")).toBe(false);
    expect(isTemporal("bigint")).toBe(false);
  });
});

describe("isCategorical", () => {
  it("returns true for string and boolean types", () => {
    expect(isCategorical("string")).toBe(true);
    expect(isCategorical("boolean")).toBe(true);
    expect(isCategorical("bool")).toBe(true);
  });

  it("returns true for char/varchar prefix types", () => {
    expect(isCategorical("char(10)")).toBe(true);
    expect(isCategorical("varchar(255)")).toBe(true);
  });

  it("is case-insensitive", () => {
    expect(isCategorical("STRING")).toBe(true);
    expect(isCategorical("BOOLEAN")).toBe(true);
  });

  it("returns false for non-categorical types", () => {
    expect(isCategorical("bigint")).toBe(false);
    expect(isCategorical("timestamp")).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// renderChart - null cases
// ---------------------------------------------------------------------------

describe("renderChart - null cases", () => {
  it("returns null for an empty result (no rows)", () => {
    const result = makeResult(
      [{ name: "category", type: "string" }, { name: "total", type: "bigint" }],
      [],
    );
    expect(renderChart(result)).toBeNull();
  });

  it("returns null when there is no numeric column", () => {
    const result = makeResult(
      [{ name: "name", type: "string" }, { name: "label", type: "string" }],
      [["a", "x"], ["b", "y"]],
    );
    expect(renderChart(result)).toBeNull();
  });

  it("returns null for a single string column with no numeric", () => {
    const result = makeResult(
      [{ name: "tag", type: "string" }],
      [["foo"]],
    );
    expect(renderChart(result)).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// renderChart - bar chart (categorical + numeric)
// ---------------------------------------------------------------------------

describe("renderChart - bar chart", () => {
  const barResult = makeResult(
    [
      { name: "category", type: "string" },
      { name: "total", type: "bigint" },
    ],
    [
      ["apples", 120],
      ["bananas", 85],
      ["cherries", 200],
    ],
  );

  it("returns an HTMLElement (div.demo-chart wrapper)", () => {
    const el = renderChart(barResult);
    expect(el).not.toBeNull();
    expect(el?.tagName.toLowerCase()).toBe("div");
    expect(el?.className).toBe("demo-chart");
  });

  it("contains a child SVG element", () => {
    const el = renderChart(barResult)!;
    const svg = el.querySelector("svg");
    expect(svg).not.toBeNull();
  });

  it("renders one rect per data row (3 bars for 3 rows)", () => {
    const el = renderChart(barResult)!;
    const rects = el.querySelectorAll("rect");
    // One rect for the white background + one per bar = 4 total
    // Background rect + 3 bars = 4
    expect(rects.length).toBe(4);
  });

  it("uses orange fill (#e25a1c) for bars", () => {
    const el = renderChart(barResult)!;
    const rects = Array.from(el.querySelectorAll("rect"));
    const barRects = rects.filter((r) => r.getAttribute("fill") === "#e25a1c");
    expect(barRects.length).toBe(3);
  });

  it("shows a title text with 'total by category'", () => {
    const el = renderChart(barResult)!;
    const texts = Array.from(el.querySelectorAll("text"));
    const titleText = texts.find((t) => t.textContent?.includes("total") && t.textContent?.includes("category"));
    expect(titleText).not.toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// renderChart - line chart (temporal + numeric)
// ---------------------------------------------------------------------------

describe("renderChart - line chart", () => {
  const lineResult = makeResult(
    [
      { name: "event_date", type: "timestamp" },
      { name: "revenue", type: "double" },
    ],
    [
      ["2024-01-01T00:00:00", 1000],
      ["2024-01-02T00:00:00", 1500],
      ["2024-01-03T00:00:00", 1200],
    ],
  );

  it("returns a non-null element for timestamp + numeric", () => {
    const el = renderChart(lineResult);
    expect(el).not.toBeNull();
  });

  it("renders a path element (the line)", () => {
    const el = renderChart(lineResult)!;
    const paths = el.querySelectorAll("path");
    // There should be at least 2 paths: area fill + line stroke
    expect(paths.length).toBeGreaterThanOrEqual(1);
  });

  it("uses blue stroke (#2272b4) for the line path", () => {
    const el = renderChart(lineResult)!;
    const paths = Array.from(el.querySelectorAll("path"));
    const linePath = paths.find((p) => p.getAttribute("stroke") === "#2272b4");
    expect(linePath).not.toBeUndefined();
  });

  it("renders circle dots for each data point (3 circles)", () => {
    const el = renderChart(lineResult)!;
    const circles = el.querySelectorAll("circle");
    expect(circles.length).toBe(3);
  });

  it("does not contain orange rects (no bars)", () => {
    const el = renderChart(lineResult)!;
    const rects = Array.from(el.querySelectorAll("rect"));
    const orangeRects = rects.filter((r) => r.getAttribute("fill") === "#e25a1c");
    expect(orangeRects.length).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// renderChart - numeric only (no category, no temporal)
// ---------------------------------------------------------------------------

describe("renderChart - numeric only", () => {
  it("renders a bar chart using row index when only numeric columns are present", () => {
    const result = makeResult(
      [{ name: "value", type: "double" }],
      [[10], [20], [30]],
    );
    const el = renderChart(result);
    expect(el).not.toBeNull();
    const rects = el!.querySelectorAll("rect");
    // background + 3 bars = 4
    expect(rects.length).toBe(4);
  });
});

// ---------------------------------------------------------------------------
// renderChart - single row
// ---------------------------------------------------------------------------

describe("renderChart - single row", () => {
  it("handles a single-row bar chart without throwing", () => {
    const result = makeResult(
      [{ name: "group", type: "string" }, { name: "count", type: "bigint" }],
      [["only", 42]],
    );
    const el = renderChart(result);
    expect(el).not.toBeNull();
    // 1 background + 1 bar = 2 rects
    expect(el!.querySelectorAll("rect").length).toBe(2);
  });

  it("handles a single-row line chart without throwing", () => {
    const result = makeResult(
      [{ name: "ts", type: "date" }, { name: "val", type: "float" }],
      [["2024-06-01", 99.9]],
    );
    const el = renderChart(result);
    expect(el).not.toBeNull();
    // Should render exactly 1 circle dot for the single point
    expect(el!.querySelectorAll("circle").length).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// renderChart - category cap (>16 rows shows "+N more" note)
// ---------------------------------------------------------------------------

describe("renderChart - category cap", () => {
  it("caps at 16 bars and shows extra note text when >16 categories", () => {
    const rows: unknown[][] = Array.from({ length: 20 }, (_, i) => [`cat${i}`, i * 10]);
    const result = makeResult(
      [{ name: "cat", type: "string" }, { name: "val", type: "int" }],
      rows,
    );
    const el = renderChart(result)!;
    // 1 background + 16 bars = 17
    const rects = el.querySelectorAll("rect");
    expect(rects.length).toBe(17);

    // Should include "+4 more" text somewhere
    const texts = Array.from(el.querySelectorAll("text"));
    const noteText = texts.find((t) => t.textContent?.includes("+4 more"));
    expect(noteText).not.toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// renderChart - null values in data
// ---------------------------------------------------------------------------

describe("renderChart - null cell values", () => {
  it("skips null-value rows (no bar rendered for null)", () => {
    const result = makeResult(
      [{ name: "label", type: "string" }, { name: "num", type: "int" }],
      [["a", 10], ["b", null], ["c", 30]],
    );
    const el = renderChart(result)!;
    // Only 2 non-null bars + 1 background rect = 3
    const orangeRects = Array.from(el.querySelectorAll("rect")).filter(
      (r) => r.getAttribute("fill") === "#e25a1c",
    );
    expect(orangeRects.length).toBe(2);
  });
});
