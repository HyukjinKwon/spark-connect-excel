// SPDX-License-Identifier: Apache-2.0
//
// typeFormat.ts — Lane F
//
// Pure, stateless helpers that translate Spark SQL type names into Excel
// number-format strings and coerce marshalled cell values into what Office.js
// `range.values` expects.
//
// Both functions are deliberately free of Office.js imports so they can be
// unit-tested with plain Vitest (no Office mock required).

// ---------------------------------------------------------------------------
// Excel serial date epoch: Excel counts days since 1899-12-30 (the "1900 date
// system" with its famous off-by-one leap-day bug preserved for Lotus compat).
// We compute the offset from the Unix epoch (1970-01-01) in days.
//
// Difference: days between 1899-12-30 and 1970-01-01
//   = 25569  (tested against Excel's own DATE(1970,1,1) = 25569)
// ---------------------------------------------------------------------------
const EXCEL_EPOCH_OFFSET_DAYS = 25569;
const MS_PER_DAY = 86_400_000;

/**
 * Map a Spark SQL type name to an Excel number-format string.
 *
 * Returns `null` when the General/"@" (text) format is appropriate, which
 * tells the caller to leave the column format unchanged (General).
 *
 * Type name grammar (from Spark):
 *   - atomic: "bigint" | "int" | "smallint" | "tinyint" | "long"
 *             "double" | "float" | "real"
 *             "decimal(p,s)" | "decimal"
 *             "date" | "timestamp" | "timestamp_ntz"
 *             "boolean" | "string" | "binary" | "void"
 *   - complex: "array<…>" | "map<…>" | "struct<…>" — treated as text
 */
export function numberFormatFor(sparkType: string): string | null {
  const t = sparkType.trim().toLowerCase();

  // Integer types → plain integer display
  if (
    t === "bigint" ||
    t === "int" ||
    t === "integer" ||
    t === "smallint" ||
    t === "tinyint" ||
    t === "long" ||
    t === "short" ||
    t === "byte"
  ) {
    return "0";
  }

  // Floating-point types → two decimal places
  if (t === "double" || t === "float" || t === "real") {
    return "0.00";
  }

  // Decimal / Numeric: extract scale and produce an appropriate format.
  // "numeric" is Spark's alias for DECIMAL (ANSI SQL synonym).
  // e.g. "decimal(10,2)" → "0.00", "numeric(18,0)" → "0", "decimal" → "0.00"
  const decimalMatch = t.match(/^(?:decimal|numeric)(?:\((\d+),\s*(\d+)\))?$/);
  if (decimalMatch) {
    const scale = decimalMatch[2] !== undefined ? parseInt(decimalMatch[2], 10) : 2;
    if (scale === 0) {
      return "0";
    }
    return "0." + "0".repeat(scale);
  }

  // Date → ISO-style date only
  if (t === "date") {
    return "yyyy-mm-dd";
  }

  // Timestamp (with or without timezone) → full datetime
  if (t === "timestamp" || t === "timestamp_ntz" || t === "timestamp_ltz") {
    return "yyyy-mm-dd hh:mm:ss";
  }

  // Boolean, string, binary, void, complex types → General / text, no format
  return null;
}

/**
 * Coerce a marshalled SparkResult cell value into a value suitable for
 * `range.values` assignment in Office.js.
 *
 * Office.js `range.values` accepts: number | string | boolean | null (empty).
 *
 * Strategy:
 *   - null → null (empty cell)
 *   - boolean → boolean (passed through)
 *   - date/timestamp columns with an ISO string value → Excel serial date
 *     number so that number-format rendering works correctly.  Falls back to
 *     the ISO string if parsing fails (safe degradation).
 *   - integer/float/decimal with a numeric JS value → number (passed through)
 *   - everything else → string via String()
 *
 * This function is pure and has no side effects.
 */
export function coerceCellValue(
  value: unknown,
  sparkType: string,
): string | number | boolean | null {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === "boolean") {
    return value;
  }

  const t = sparkType.trim().toLowerCase();
  const isDateType = t === "date";
  const isTimestampType = t === "timestamp" || t === "timestamp_ntz" || t === "timestamp_ltz";

  if ((isDateType || isTimestampType) && typeof value === "string") {
    return isoToExcelSerial(value, isDateType);
  }

  if (typeof value === "number") {
    // Pass numeric values straight through (handles int, float, decimal as
    // marshalled by Python's json.dumps which preserves JS numeric types).
    return value;
  }

  // For everything else (string cells, complex types rendered as JSON, etc.)
  return String(value);
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Convert an ISO-8601 date or datetime string to an Excel serial date number.
 *
 * Excel serial date = number of days since 1899-12-30 (UTC midnight for date-
 * only, or with fractional days for datetime).
 *
 * Returns the original string if the date is invalid so the cell at least
 * shows something human-readable rather than #VALUE!.
 */
export function isoToExcelSerial(iso: string, dateOnly: boolean): number | string {
  // Attempt to parse.  For date-only ("2024-03-15") we parse as UTC midnight
  // to avoid timezone-induced day shifts.  For timestamps we use the ISO
  // string directly (Python marshals them with a timezone offset or as UTC).
  let ms: number;

  if (dateOnly) {
    // Ensure we interpret the date in UTC regardless of local timezone by
    // appending "T00:00:00Z" if there is no time component.
    const normalized = /^\d{4}-\d{2}-\d{2}$/.test(iso) ? iso + "T00:00:00Z" : iso;
    ms = Date.parse(normalized);
  } else {
    // Timestamps: Python emits ISO strings; if there is no timezone info we
    // treat them as UTC (consistent with Spark's internal representation).
    // Timezone indicators: trailing "Z", or an offset like "+05:30" / "-07:00".
    const hasZoneInfo = /Z$|[+-]\d{2}:\d{2}$/.test(iso.trim());
    const normalized = hasZoneInfo ? iso : iso + "Z";
    ms = Date.parse(normalized);
  }

  if (isNaN(ms)) {
    // Graceful fallback: return the raw string so the cell is still readable.
    return iso;
  }

  // Convert milliseconds from Unix epoch to Excel serial days.
  const unixDays = ms / MS_PER_DAY;
  return unixDays + EXCEL_EPOCH_OFFSET_DAYS;
}
