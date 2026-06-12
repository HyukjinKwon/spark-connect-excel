// SPDX-License-Identifier: Apache-2.0
//
// rangeWriter.ts — Lane F
//
// Writes a SparkResult into an Excel worksheet range via Office.js.
//
// Conventions (shared across excel/* lanes):
//   * Every public mutating function is self-contained — it calls
//     Excel.run(async (ctx) => { … }) itself and syncs.
//   * Callers invoke them sequentially.
//
// Layout (rows are 0-indexed in code; addresses use 1-indexed Excel notation):
//
//   [row A]  optional truncation banner   ← only present when result.truncated
//   [row B]  header row                   ← always
//   [row C…] body rows                    ← 0 rows if result.rows is empty
//
// When result.truncated is true and anchorAddress points to row 1, the banner
// is written ONE ROW BELOW the header (between header and body) rather than
// above it, because there is no room above row 1.  This choice is documented
// in docs/findings-lane-F.md.

import type { SparkResult } from "../seam.js";
import { numberFormatFor, coerceCellValue } from "./typeFormat.js";

/** Describes the ranges written by writeResult. */
export interface WriteResultInfo {
  /** Name of the worksheet that received the data. */
  sheetName: string;
  /** Address of the header row (e.g. "Sheet1!A1:D1"). */
  headerRangeAddress: string;
  /** Address of the body rows, or empty string when rowCount === 0. */
  bodyRangeAddress: string;
  /**
   * Address of header + body together (the "data range").
   * Other lanes (G binding, H charting) anchor their work here.
   */
  dataRangeAddress: string;
  /** Number of data rows written (excludes header and banner rows). */
  rowCount: number;
  /** Number of columns. */
  colCount: number;
  /** Mirrors SparkResult.truncated — true if the row cap clipped the result. */
  truncated: boolean;
}

// ---------------------------------------------------------------------------
// Styling constants
// ---------------------------------------------------------------------------

/** Background fill for the header row (medium blue). */
const HEADER_BG_COLOR = "#1F4E79";
/** Font color for the header row text (white). */
const HEADER_FONT_COLOR = "#FFFFFF";
/** Background fill for the truncation banner (amber). */
const BANNER_BG_COLOR = "#FFC000";
/** Font color for the truncation banner text. */
const BANNER_FONT_COLOR = "#000000";

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Write a SparkResult into an Excel worksheet range.
 *
 * @param result  The query result from SparkBridge.runSQL().
 * @param opts    Optional placement options.
 *   - anchorAddress  Top-left cell address (e.g. "A1").  Defaults to the
 *                    currently selected cell.
 *   - sheetName      Worksheet name.  Defaults to the active sheet.
 *
 * @returns WriteResultInfo describing the written ranges for downstream
 *          consumers (binding, charting).
 */
export async function writeResult(
  result: SparkResult,
  opts?: { anchorAddress?: string; sheetName?: string },
): Promise<WriteResultInfo> {
  return Excel.run(async (ctx) => {
    // ------------------------------------------------------------------
    // 1. Resolve worksheet and anchor cell.
    // ------------------------------------------------------------------
    const sheet =
      opts?.sheetName != null
        ? ctx.workbook.worksheets.getItem(opts.sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    // Determine anchor: use provided address or fall back to selected cell.
    let anchorCell: Excel.Range;
    if (opts?.anchorAddress != null) {
      anchorCell = sheet.getRange(opts.anchorAddress);
    } else {
      anchorCell = ctx.workbook.getSelectedRange();
    }

    // Load anchor address so we can check whether we have room above.
    anchorCell.load("address,rowIndex");
    await ctx.sync();

    const sheetName = sheet.name;
    const colCount = result.schema.length;
    const rowCount = result.rows.length;

    // ------------------------------------------------------------------
    // 2. Decide banner placement.
    //
    // Preferred: banner goes ONE ROW ABOVE the header (anchorCell.rowIndex
    // is 0-based; rowIndex === 0 means we are on row 1 — no room above).
    //
    // Fallback when anchored at row 1: banner is written ONE ROW BELOW the
    // header (between header and body).  This is safe because even if
    // result.rows is empty we will produce a "(0 rows)" note that occupies
    // that slot, so there is always at least one body row of space.
    // ------------------------------------------------------------------
    const anchorRowIndex: number = anchorCell.rowIndex; // 0-based
    const bannerAbove = result.truncated && anchorRowIndex > 0;
    const bannerBelow = result.truncated && anchorRowIndex === 0;

    // The header row index (0-based in worksheet terms).
    const headerRowOffset = bannerAbove ? 1 : 0; // relative to anchorCell row
    const headerAbsRow = anchorRowIndex + headerRowOffset;

    // ------------------------------------------------------------------
    // 3. Build the 2D values array for header + body in one allocation.
    //
    // We build: headerValues (1 × colCount) and bodyValues (rowCount × colCount).
    // We also build number-format arrays.
    // ------------------------------------------------------------------
    const colNames = result.schema.map((c) => c.name);
    const colTypes = result.schema.map((c) => c.type);

    // Header row values.
    const headerValues: (string | number | boolean | null)[][] = [colNames];

    // Body row values — coerce each cell through typeFormat.
    let bodyValues: (string | number | boolean | null)[][];
    if (rowCount > 0) {
      bodyValues = result.rows.map((row) =>
        colTypes.map((type, ci) => coerceCellValue(row[ci], type)),
      );
    } else {
      // Empty result: write a single "(0 rows)" note spanning all columns.
      // Office.js doesn't support true merge in range.values, so we write
      // the text in the first cell and leave the rest empty.
      bodyValues = [colTypes.map((_, ci) => (ci === 0 ? "(0 rows)" : null))];
    }

    // Number-format row for body columns (one format string per column,
    // applied as a 1-row template that Office.js repeats for all body rows).
    const bodyFormats: string[][] = [colTypes.map((type) => numberFormatFor(type) ?? "General")];

    // ------------------------------------------------------------------
    // 4. Write truncation banner ABOVE the header (preferred path).
    // ------------------------------------------------------------------
    if (bannerAbove) {
      // anchorRowIndex > 0, so the banner goes at the anchor cell's row.
      const bannerRange = anchorCell.getResizedRange(0, colCount - 1);
      bannerRange.merge(true); // merge across all columns for readability
      bannerRange.values = [[`⚠️ Showing first ${result.rowCount} rows (result truncated)`]];
      bannerRange.format.fill.color = BANNER_BG_COLOR;
      bannerRange.format.font.color = BANNER_FONT_COLOR;
      bannerRange.format.font.bold = true;
    }

    // ------------------------------------------------------------------
    // 5. Write the header row.
    // ------------------------------------------------------------------
    const headerTopCell = anchorCell.getOffsetRange(headerRowOffset, 0);
    const headerRange = headerTopCell.getResizedRange(0, colCount - 1);
    headerRange.values = headerValues as Excel.Range["values"];
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = HEADER_BG_COLOR;
    headerRange.format.font.color = HEADER_FONT_COLOR;

    // Load the header address for the return value.
    headerRange.load("address");

    // ------------------------------------------------------------------
    // 6. Write the body rows.
    // ------------------------------------------------------------------
    // Body starts one row below the header, regardless of banner placement.
    const bodyRowsBelowHeader = bannerBelow
      ? 2 // header + banner-below + body; body is 2 rows below header cell
      : 1; // normal: body is 1 row below header

    const bodyTopCell = headerTopCell.getOffsetRange(bodyRowsBelowHeader, 0);

    // Actual number of rows we write (0-rows case still writes 1 note row).
    const bodyWriteRowCount = Math.max(rowCount, 1);
    const bodyRange = bodyTopCell.getResizedRange(bodyWriteRowCount - 1, colCount - 1);
    bodyRange.values = bodyValues as Excel.Range["values"];

    // Apply number formats — Office.js accepts a 2D format array whose shape
    // must match the range dimensions.  We build one format row and replicate
    // it for every data row.
    const bodyFormatsExpanded: string[][] = Array.from(
      { length: bodyWriteRowCount },
      () => bodyFormats[0],
    );
    bodyRange.numberFormat = bodyFormatsExpanded;
    bodyRange.load("address");

    // ------------------------------------------------------------------
    // 7. Write truncation banner BELOW the header (fallback path: row 1).
    // ------------------------------------------------------------------
    if (bannerBelow) {
      // bannerBelow means anchorRowIndex === 0 and result.truncated.
      // Banner sits between header and body (row 2 in 1-based terms).
      const bannerBelowCell = headerTopCell.getOffsetRange(1, 0);
      const bannerBelowRange = bannerBelowCell.getResizedRange(0, colCount - 1);
      bannerBelowRange.merge(true);
      bannerBelowRange.values = [
        [
          `⚠️ Showing first ${result.rowCount} rows (result truncated) — anchored at row 1, banner placed below header`,
        ],
      ];
      bannerBelowRange.format.fill.color = BANNER_BG_COLOR;
      bannerBelowRange.format.font.color = BANNER_FONT_COLOR;
      bannerBelowRange.format.font.bold = true;
    }

    // ------------------------------------------------------------------
    // 8. Freeze the header row (freeze pane below the header row).
    // ------------------------------------------------------------------
    // freezePanes.freezeRows(n) freezes the first n rows (1-based count).
    // We want to freeze everything up through the header row.
    // headerAbsRow is 0-based, so the header is row (headerAbsRow + 1) in
    // 1-based Excel terms; we freeze that many rows.
    sheet.freezePanes.freezeRows(headerAbsRow + 1);

    // ------------------------------------------------------------------
    // 9. Autofit columns across the whole data range.
    // ------------------------------------------------------------------
    // Build a range covering all written columns from header to last body row.
    const totalBodyRows = bannerBelow ? bodyWriteRowCount + 1 : bodyWriteRowCount;
    const dataRange = headerTopCell.getResizedRange(totalBodyRows, colCount - 1);
    dataRange.format.autofitColumns();
    dataRange.load("address");

    // Final sync to populate all loaded properties.
    await ctx.sync();

    // ------------------------------------------------------------------
    // 10. Compute return addresses.
    // ------------------------------------------------------------------
    const headerAddr = stripSheetPrefix(headerRange.address);
    const bodyAddr = rowCount > 0 ? stripSheetPrefix(bodyRange.address) : "";

    // dataRangeAddress = header + body (the range other lanes bind charts to).
    // We use the already-loaded dataRange.address for this.
    const dataAddr = stripSheetPrefix(dataRange.address);

    return {
      sheetName,
      headerRangeAddress: headerAddr,
      bodyRangeAddress: bodyAddr,
      dataRangeAddress: dataAddr,
      rowCount,
      colCount,
      truncated: result.truncated,
    };
  });
}

// ---------------------------------------------------------------------------
// Private helpers
// ---------------------------------------------------------------------------

/**
 * Strip the sheet prefix from a fully-qualified address.
 *
 * Office.js returns addresses like "Sheet1!A1:D5" or "'My Sheet'!A1:D5".
 * Other lanes (G, H) typically want just the local part ("A1:D5") or the full
 * form — we store the full form so the sheet is unambiguous.
 *
 * Actually, let's keep the full address (sheet-qualified) for unambiguity.
 * This helper is a no-op but documents the intent.
 */
function stripSheetPrefix(address: string): string {
  // Keep the full sheet-qualified address for safety.
  return address;
}
