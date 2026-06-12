// SPDX-License-Identifier: Apache-2.0
//
// chart.ts — Lane H: native Excel chart insertion from a Spark query result.
//
// Public API (consumed by Lane E):
//   inferChartType(schema)         — pure, runtime-testable, returns ChartKind
//   toExcelChartType(kind)         — maps ChartKind → Excel.ChartType enum value
//   insertChart(info, schema, opts) — Excel.run wrapper; adds a chart to the sheet
//
// Design note — ChartKind / runtime split
// ----------------------------------------
// Excel.ChartType is a TypeScript ambient enum from @types/office-js. Its numeric
// values (e.g. Excel.ChartType.Line = 65) are NOT available at runtime in jsdom or
// any non-Office environment, because office-js is never actually loaded in unit
// tests. Returning Excel.ChartType directly from inferChartType would make the pure
// heuristic untestable without a live Office host.
//
// Solution: inferChartType returns a `ChartKind` string-literal union. A separate
// toExcelChartType() maps that string to the Excel.ChartType enum value at call
// sites that already have the Office host available (i.e., inside Excel.run).
// Tests verify inferChartType against ChartKind strings — no Office.js needed.

import type { ColumnMeta } from "../seam.js";

// ---------------------------------------------------------------------------
// WriteResultInfo — shape promised by Lane F (rangeWriter.ts). Coded against
// the canonical shape; import the real type once Lane F ships.
// ---------------------------------------------------------------------------

export interface WriteResultInfo {
  sheetName: string;
  headerRangeAddress: string;
  bodyRangeAddress: string;
  /** Combined header + body range (used as chart data source). */
  dataRangeAddress: string;
  rowCount: number;
  colCount: number;
  truncated: boolean;
}

// ---------------------------------------------------------------------------
// ChartKind — runtime-safe string-literal union (see design note above)
// ---------------------------------------------------------------------------

/**
 * A runtime-safe alternative to Excel.ChartType. inferChartType returns this;
 * toExcelChartType maps it to the Office.js enum value.
 *
 * Supported kinds and their Excel.ChartType equivalents:
 *   "line"             → Excel.ChartType.Line           (65)
 *   "columnClustered"  → Excel.ChartType.ColumnClustered (51)
 *   "columnStacked"    → Excel.ChartType.ColumnStacked  (52)
 *   "xyScatter"        → Excel.ChartType.XYScatter      (-4169)
 */
export type ChartKind = "line" | "columnClustered" | "columnStacked" | "xyScatter";

// ---------------------------------------------------------------------------
// Spark type classifiers (pure helpers)
// ---------------------------------------------------------------------------

/** Numeric Spark SQL type names (exact match or prefix for decimal/numeric). */
function isNumeric(type: string): boolean {
  const t = type.toLowerCase();
  return (
    t === "bigint" ||
    t === "int" ||
    t === "integer" ||
    t === "smallint" ||
    t === "tinyint" ||
    t === "double" ||
    t === "float" ||
    t === "real" ||
    t.startsWith("decimal") ||
    t.startsWith("numeric")
  );
}

/** Temporal Spark SQL type names. */
function isTemporal(type: string): boolean {
  const t = type.toLowerCase();
  return t === "date" || t === "timestamp" || t === "timestamp_ntz" || t === "timestamp_ltz";
}

/** Categorical (string-like or boolean) Spark SQL type names. */
function isCategorical(type: string): boolean {
  const t = type.toLowerCase();
  return (
    t === "string" ||
    t === "boolean" ||
    t === "bool" ||
    t.startsWith("char") ||
    t.startsWith("varchar")
  );
}

// ---------------------------------------------------------------------------
// inferChartType — pure heuristic, no Office.js dependency
// ---------------------------------------------------------------------------

/**
 * Infer a suitable chart kind from the result schema.
 *
 * Heuristic rules (evaluated in priority order):
 *
 * 1. Temporal + ≥1 numeric           → "line"
 *    (time-series; the temporal column drives the X axis)
 *
 * 2. Exactly 1 categorical + 1 numeric → "columnClustered"
 *    (simple bar chart — one dimension, one measure)
 *
 * 3. 1 categorical + multiple numeric  → "columnClustered"
 *    (multi-measure bar; ColumnStacked is an alternative but clustered is
 *    safer as a default since it preserves individual series readability.
 *    Callers can override via opts.chartType in insertChart.)
 *
 * 4. ≥2 numeric + no categorical       → "xyScatter"
 *    (scatter / correlation; first numeric = X, remainder = Y series)
 *
 * 5. Default                           → "columnClustered"
 *    (covers mixed or unknown schemas)
 */
export function inferChartType(schema: ColumnMeta[]): ChartKind {
  const numericCols = schema.filter((c) => isNumeric(c.type));
  const temporalCols = schema.filter((c) => isTemporal(c.type));
  const categoricalCols = schema.filter((c) => isCategorical(c.type));

  // Rule 1: time-series
  if (temporalCols.length >= 1 && numericCols.length >= 1) {
    return "line";
  }

  // Rule 2 & 3: categorical + numeric(s)
  if (categoricalCols.length === 1 && numericCols.length >= 1) {
    return "columnClustered";
  }

  // Rule 4: scatter (two or more numerics, no good category axis)
  if (numericCols.length >= 2 && categoricalCols.length === 0) {
    return "xyScatter";
  }

  // Rule 5: default
  return "columnClustered";
}

// ---------------------------------------------------------------------------
// toExcelChartType — ChartKind → Excel.ChartType (requires Office host)
// ---------------------------------------------------------------------------

/**
 * Map a ChartKind string to the corresponding Excel.ChartType enum value.
 * Must only be called inside an Excel.run callback (Office.js must be loaded).
 */
export function toExcelChartType(kind: ChartKind): Excel.ChartType {
  switch (kind) {
    case "line":
      return Excel.ChartType.line;
    case "columnClustered":
      return Excel.ChartType.columnClustered;
    case "columnStacked":
      return Excel.ChartType.columnStacked;
    case "xyScatter":
      return Excel.ChartType.xyscatter;
  }
}

// ---------------------------------------------------------------------------
// insertChart — public async entry point (self-contained Excel.run)
// ---------------------------------------------------------------------------

export interface InsertChartOptions {
  /** Override the inferred chart type. */
  chartType?: Excel.ChartType;
  /** Override the derived chart title. */
  title?: string;
}

export interface InsertChartResult {
  /** The Office-assigned name of the created chart object. */
  chartName: string;
}

/**
 * Insert a native Excel chart bound to the result range produced by Lane F's
 * writeResult(). Positions the chart to the right of the data range.
 *
 * @param info   WriteResultInfo from Lane F's rangeWriter.
 * @param schema Column metadata (used for chart-type inference and title).
 * @param opts   Optional overrides for chart type and title.
 *
 * @throws Error("Not enough data to chart") when rowCount is 0 or the schema
 *         has fewer than 2 columns (Excel requires at least one category and
 *         one value series, or two value series for scatter).
 */
export async function insertChart(
  info: WriteResultInfo,
  schema: ColumnMeta[],
  opts?: InsertChartOptions,
): Promise<InsertChartResult> {
  // Guard: must have data rows and a meaningful schema
  if (info.rowCount === 0) {
    throw new Error("Not enough data to chart: result set is empty (rowCount = 0).");
  }
  if (schema.length < 2) {
    throw new Error(
      "Not enough data to chart: schema has fewer than 2 columns " +
        "(need at least one category/value and one measure).",
    );
  }

  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(info.sheetName);

    // Resolve the data range (header + body, as written by Lane F)
    const dataRange = worksheet.getRange(info.dataRangeAddress);

    // Determine chart type
    const kind = inferChartType(schema);
    const excelChartType = opts?.chartType ?? toExcelChartType(kind);

    // Add the chart; Excel auto-selects series from the data range
    const chart = worksheet.charts.add(excelChartType, dataRange, Excel.ChartSeriesBy.auto);

    // Set chart title
    const title = opts?.title ?? deriveTitle(schema);
    chart.title.text = title;
    chart.title.visible = true;

    // Position the chart to the right of the data range.
    // We use setPosition with a top-left cell just past the data columns and a
    // bottom-right cell 15 rows down × 8 columns wide, giving a reasonable
    // default size that doesn't overlap the data.
    const topLeftAddress = rightOfData(info.dataRangeAddress, info.colCount);
    const bottomRightAddress = shiftAddress(topLeftAddress, 15, 8);
    chart.setPosition(topLeftAddress, bottomRightAddress);

    // Load the assigned name so we can return it
    chart.load("name");
    await context.sync();

    return { chartName: chart.name };
  });
}

// ---------------------------------------------------------------------------
// Private helpers
// ---------------------------------------------------------------------------

/**
 * Derive a human-readable chart title from the schema columns.
 * Format: "<measure(s)> by <dimension>" or "<col1> vs <col2>".
 */
function deriveTitle(schema: ColumnMeta[]): string {
  const temporalCols = schema.filter((c) => isTemporal(c.type));
  const categoricalCols = schema.filter((c) => isCategorical(c.type));
  const numericCols = schema.filter((c) => isNumeric(c.type));

  if (temporalCols.length >= 1 && numericCols.length >= 1) {
    const measures = numericCols.map((c) => c.name).join(", ");
    return `${measures} over ${temporalCols[0].name}`;
  }
  if (categoricalCols.length >= 1 && numericCols.length >= 1) {
    const measures = numericCols.map((c) => c.name).join(", ");
    return `${measures} by ${categoricalCols[0].name}`;
  }
  if (numericCols.length >= 2) {
    return `${numericCols[0].name} vs ${numericCols[1].name}`;
  }
  // Generic fallback: join first two column names
  return schema
    .slice(0, 2)
    .map((c) => c.name)
    .join(" & ");
}

/**
 * Given a range address like "A1:C10", return the address of the cell
 * immediately to the right of the rightmost column in the top row.
 *
 * E.g. dataRangeAddress="A1:C10", colCount=3 → "D1"
 *
 * This is a simple letter-arithmetic helper; it handles single-letter
 * column names (A–Z) and double-letter names (AA–ZZ) which covers the
 * practical range for Spark result sets within the 10k-row cap.
 */
/** @internal Exported for unit testing only. */
export function rightOfData(dataRangeAddress: string, colCount: number): string {
  // Extract the top-left cell of the range (e.g. "A1" from "A1:C10" or
  // "Sheet1!A1:C10"). Strip any sheet prefix (everything up to and including
  // the last "!") so the regex works on the bare cell address.
  const topLeftFull = dataRangeAddress.split(":")[0];
  const topLeft = topLeftFull.split("!").pop() ?? topLeftFull;
  // Split into column letters and row digits
  const match = /^([A-Z]+)(\d+)$/.exec(topLeft);
  if (!match) {
    // Fallback: place at a fixed offset if address parsing fails
    return "J1";
  }
  const colLetters = match[1];
  const rowDigit = match[2];
  // Advance the column by colCount positions
  const newColLetters = columnLettersFromIndex(columnIndexFromLetters(colLetters) + colCount);
  return `${newColLetters}${rowDigit}`;
}

/**
 * Shift an address right by `cols` and down by `rows`.
 * E.g. shiftAddress("D1", 15, 8) → "L16"
 */
function shiftAddress(address: string, rows: number, cols: number): string {
  const match = /^([A-Z]+)(\d+)$/.exec(address);
  if (!match) return address;
  const newCol = columnLettersFromIndex(columnIndexFromLetters(match[1]) + cols);
  const newRow = parseInt(match[2], 10) + rows;
  return `${newCol}${newRow}`;
}

/** Convert column letters (A, B, …, Z, AA, …) to a 0-based index. */
function columnIndexFromLetters(letters: string): number {
  let index = 0;
  for (const ch of letters) {
    index = index * 26 + (ch.charCodeAt(0) - 64);
  }
  return index; // 1-based: A=1, B=2, …
}

/** Convert a 1-based column index back to letters. */
function columnLettersFromIndex(index: number): string {
  let letters = "";
  while (index > 0) {
    const rem = (index - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    index = Math.floor((index - 1) / 26);
  }
  return letters;
}
