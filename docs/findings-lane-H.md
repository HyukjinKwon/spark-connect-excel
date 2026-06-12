<!-- SPDX-License-Identifier: Apache-2.0 -->

# Lane H findings — Native Excel chart insertion

## 1. Files delivered by Lane H

| File | Role |
|------|------|
| `src/excel/chart.ts` | `inferChartType`, `toExcelChartType`, `insertChart`, `WriteResultInfo`, `ChartKind` |
| `tests/unit/chart.test.ts` | Vitest unit tests for `inferChartType` (pure, no Office.js) |
| `docs/findings-lane-H.md` | This file |

---

## 2. Chart-type heuristic (`inferChartType`)

Rules are evaluated in priority order. The first matching rule wins.

| Priority | Condition | ChartKind returned | Rationale |
|----------|-----------|--------------------|-----------|
| 1 | ≥1 temporal column **and** ≥1 numeric | `"line"` | Time-series data reads best as a line; temporal col drives X axis |
| 2 | Exactly 1 categorical + 1 numeric | `"columnClustered"` | Classic bar chart — one dimension, one measure |
| 3 | 1 categorical + ≥2 numerics | `"columnClustered"` | Multi-measure bar; clustered over stacked because it preserves individual series magnitude readability. Callers can override via `opts.chartType`. |
| 4 | ≥2 numerics and 0 categoricals | `"xyScatter"` | Correlation / scatter plot when there is no categorical axis |
| 5 | Anything else | `"columnClustered"` | Safe default; always produces a valid chart |

### Spark type classifiers

```
isNumeric     → bigint | int | integer | smallint | tinyint | double | float |
                real | decimal(...) | numeric(...)
isTemporal    → date | timestamp | timestamp_ntz | timestamp_ltz
isCategorical → string | boolean | bool | char(...) | varchar(...)
```

Matching is **case-insensitive** (`.toLowerCase()`); prefix matching handles
parameterised types (`decimal(18,2)`, `varchar(64)`, etc.).

---

## 3. ChartKind / runtime split

### Problem

`Excel.ChartType` is a TypeScript **ambient enum** supplied by `@types/office-js`.
Its numeric member values (e.g. `Excel.ChartType.Line = 65`) exist only inside a
live Office host. In jsdom (vitest's test environment) `office-js` is never loaded,
so referencing `Excel.ChartType.Line` at runtime throws `ReferenceError: Excel is
not defined`.

### Solution

`inferChartType` returns a **`ChartKind` string-literal union** (`"line"` |
`"columnClustered"` | `"columnStacked"` | `"xyScatter"`). A separate mapper,
`toExcelChartType(kind): Excel.ChartType`, converts that string to the Office enum
value. `toExcelChartType` is only called inside an `Excel.run` callback where the
Office host is guaranteed to be present.

```
inferChartType(schema)  -->  ChartKind (string)     <- pure, testable in jsdom
toExcelChartType(kind)  -->  Excel.ChartType         <- Office.js required
worksheet.charts.add(…)                              <- Excel.run only
```

### Impact on Lane E

Lane E calls `insertChart(info, schema, opts?)`. It never needs to import `ChartKind`
or `toExcelChartType` directly. Lane E only imports `InsertChartOptions` and
`InsertChartResult` if it needs to type the call site.

---

## 4. `insertChart` behaviour

```ts
export async function insertChart(
  info: WriteResultInfo,
  schema: ColumnMeta[],
  opts?: InsertChartOptions,
): Promise<InsertChartResult>
```

- **Guard:** throws `Error("Not enough data to chart: …")` if `info.rowCount === 0`
  or `schema.length < 2`.
- **Data range:** uses `info.dataRangeAddress` (header + body, as written by Lane F).
  `Excel.ChartSeriesBy.auto` lets Excel pick row vs. column series orientation.
- **Title:** derived from schema column names if `opts.title` is not provided.
- **Position:** placed to the right of the data range using `chart.setPosition`.
  Top-left = one column past the data's rightmost column, row 1; spans 15 rows x 8 cols.
- **Returns:** `{ chartName: string }` — the Office-assigned chart object name.

---

## 5. WriteResultInfo dependency (Lane F)

`chart.ts` declares its own `WriteResultInfo` interface matching the shape promised
by Lane F's `rangeWriter.ts`. Once Lane F ships, replace the local declaration with:

```ts
import type { WriteResultInfo } from "./rangeWriter.js";
```

No functional changes required — the shapes are identical by contract.

---

## 6. Column-letter arithmetic

`insertChart` positions the chart using A-Z / AA-ZZ column-letter helpers covering
columns A-ZZ (702 columns), far exceeding the 10k-row-cap result sets this add-in
targets. No external dependency needed.
