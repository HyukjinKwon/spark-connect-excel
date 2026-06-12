<!-- SPDX-License-Identifier: Apache-2.0 -->

# Lane F findings - Range writer & type formatting

**Date:** 2026-06-12

## 1. Excel serial date strategy

Excel stores dates and datetimes as serial numbers (floating-point count of days
since **1899-12-30** in the "1900 date system").  The Lotus 1-2-3 leap-day bug
means Excel internally believes 1900 was a leap year; this shifts serial numbers
by 1 for all dates after 1900-02-28, but has been stable since Lotus 1-2-3 and
is intentionally preserved.

**Conversion formula:**

```
excelSerial = (unix_ms / 86_400_000) + 25569
```

Where `25569` is the number of days between 1899-12-30 and the Unix epoch
(1970-01-01), verified against Excel's own `DATE(1970,1,1) = 25569`.

### Date-only values (`date` type)

Python marshals `date` objects as ISO-8601 strings without a time component
(e.g. `"2024-03-15"`).  We append `T00:00:00Z` before `Date.parse()` to force
UTC interpretation and avoid local-timezone day shifts (a user in UTC+9 would
otherwise see `"2024-03-15"` parsed as the start of March 14 UTC, shifting the
serial by -1).

### Timestamp values (`timestamp`, `timestamp_ntz`, `timestamp_ltz`)

Python marshals timestamps as ISO-8601 strings, optionally with a timezone
offset.  If no timezone info is present (common for `timestamp_ntz`), we append
`"Z"` to treat the value as UTC, consistent with Spark's internal representation.
Fractional days in the serial encode the time-of-day component, which the Excel
number format `"yyyy-mm-dd hh:mm:ss"` renders correctly.

### Fallback

If `Date.parse()` returns `NaN` (malformed string), `isoToExcelSerial` returns
the raw ISO string instead of a number.  This means the cell shows the string
rather than triggering an Office.js type error or writing `NaN` into the grid.

---

## 2. Truncation banner placement

The spec (API_CONTRACT.md section 4) calls for a "clearly-formatted note one row ABOVE
the header."

### Preferred path (anchor row > 1)

When the anchor is not in row 1, the banner occupies the anchor cell's row and
the header is written one row below it.  The banner is merged across all
columns, filled amber (`#FFC000`), and prefixed with a warning emoji.

### Fallback path (anchor at row 1)

When `anchorAddress` resolves to row 1 there is no room above.  In this case
the banner is written **between the header row and the body rows** (one row
below the header).  The banner text notes this placement.  Body data rows begin
two rows below the header in this case.

`bodyRangeAddress` in `WriteResultInfo` still points to the actual data rows
(skipping the banner), so Lane G (binding/refresh) and Lane H (charting) are
not affected.

---

## 3. Empty-result handling

When `result.rows` is empty, `writeResult` writes the header row and a single
body row containing `"(0 rows)"` in the first cell.  Returns `rowCount: 0` and
`bodyRangeAddress: ""`.

---

## 4. Files delivered by Lane F

| File | Role |
|------|------|
| `src/excel/typeFormat.ts` | `numberFormatFor` + `coerceCellValue` + `isoToExcelSerial` |
| `src/excel/rangeWriter.ts` | `writeResult` + `WriteResultInfo` |
| `tests/unit/typeFormat.test.ts` | Vitest unit tests for the pure helpers |
| `docs/findings-lane-F.md` | This file |
