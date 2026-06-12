<!-- SPDX-License-Identifier: Apache-2.0 -->

# Usage Guide

## Overview

spark-connect-excel adds a **Spark SQL** button to the Excel ribbon (Home tab →
Spark Connect group). Clicking it opens the task pane where you can query your
Spark Connect cluster and land the result directly in a worksheet range.

---

## Step 1 — Connect to your Spark cluster

Open the task pane (Home → Spark SQL). In the **Connection** section:

| Field | Description |
|-------|-------------|
| Host | Spark Connect / Envoy hostname (e.g. `localhost` or `spark.example.com`) |
| Port | Envoy grpc-web port (dev: `8081`, prod: `8443`) |
| TLS | Enable for HTTPS/TLS endpoints (required for production) |
| Bearer Token | Optional. Used by the Envoy proxy for authentication (see `docs/security.md`). Never stored in the spreadsheet. |

Click **Connect**. The status indicator will show:

- `Connecting…` — Pyodide is booting (first time only; ~10–30s)
- `Connected` — the Spark session is live
- An error message if the endpoint is unreachable

> **Tip:** The Pyodide cold-start happens once per browser session. Subsequent
> connects in the same session are fast.

<!-- Screenshot placeholder: Connection form with status "Connected" -->

---

## Step 2 — Write a SQL query

In the **Query** section:

- Enter your Spark SQL in the text area.
- Set the **Row cap** (default 10,000). Results are capped at this number of
  rows to protect the worksheet; a truncation banner appears if the result was
  clipped.
- Choose the **Destination** cell (e.g. `A1` on `Sheet1`) where the result will
  be anchored.

Examples:

```sql
-- Simple range
SELECT * FROM range(1, 11) AS t

-- Aggregation (push computation to Spark)
SELECT region, SUM(revenue) AS total_revenue
FROM sales
GROUP BY region
ORDER BY total_revenue DESC
LIMIT 20

-- Join
SELECT o.order_id, c.name, o.amount
FROM orders o
JOIN customers c ON o.customer_id = c.id
WHERE o.status = 'shipped'
```

<!-- Screenshot placeholder: SQL editor with a query entered -->

---

## Step 3 — Run the query

Click **Run**. The add-in:

1. Sends the SQL to the Spark Connect server via the Envoy proxy.
2. Waits for the result (shows a spinner during execution).
3. Writes a **header row** + **typed data rows** to the destination range:
   - Integer columns → Excel integer format (`0`)
   - Float/decimal columns → two-decimal format (`0.00` or `0.0000` based on scale)
   - Date columns → `yyyy-mm-dd`
   - Timestamp columns → `yyyy-mm-dd hh:mm:ss`
   - Boolean/string columns → General (text)
4. If the result was truncated a **truncation banner** (amber background) appears
   above the header row.

<!-- Screenshot placeholder: Result in worksheet — header row + data rows -->

---

## Step 4 — Refresh

The query and destination are **bound** to the range. To re-execute the same
query and update the range with fresh data:

- Click **Refresh** in the task pane, or
- (Future) use Ribbon → Refresh All for all bound ranges in the workbook.

The binding stores the query text, endpoint host, row cap, and destination in
the workbook's document settings (not in cells and never the bearer token — see
`docs/security.md`).

<!-- Screenshot placeholder: Refresh button in the task pane -->

---

## Step 5 — Insert a chart

After a successful run, click **Insert Chart**. The add-in:

1. Inspects the result schema and automatically chooses a chart type:
   - Temporal column + numeric column(s) → **Line chart** (time series)
   - Categorical column + numeric column(s) → **Clustered column chart**
   - Two or more numeric columns, no category → **XY Scatter**
   - Default → **Clustered column**
2. Creates a native Excel chart bound to the result range.
3. Positions the chart to the right of the data.

You can resize or reformat the chart using standard Excel tools.

<!-- Screenshot placeholder: Excel chart inserted next to data range -->

---

## Tips

### Large results

Keep the row cap at or below 10,000 for responsive results. Aggregating in Spark
(`GROUP BY`, `LIMIT`, window functions) is always better than pulling raw rows.

### Refreshing after workbook share

Bearer tokens are stored in `OfficeRuntime.storage` (roaming, per-user) and are
**not** included in the workbook file. If you share the workbook, the recipient
will need to enter their own credentials and click Connect before refreshing.

### Column types

The add-in uses Spark's schema (from `df.schema.fields`) for type inference, not
the pandas output types. This means `decimal(10,2)` produces a 2-decimal Excel
format even if the column is represented as a float64 in pandas.

### Timestamps and timezones

Timestamps without timezone info (`timestamp_ntz`) are treated as UTC. Timestamps
with timezone info use the offset embedded in the ISO-8601 string produced by
the Python runtime. Excel displays the resulting serial date according to your
local timezone settings.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| "Not connected — call connect() first" | `connect()` was not called yet | Click Connect in the task pane |
| "Spark engine starting…" for > 60s | Pyodide cold-start on slow connection | Wait; Pyodide is ~20 MB |
| HTTP 401 from Envoy | Missing or wrong bearer token | Enter token in the connection form |
| `crossOriginIsolated === false` | COI headers not served | Check the deploy/ stack or dev server config |
| Chart missing data | Empty result set | Verify the SQL returns rows; check row cap |
| Truncation banner | Result exceeds row cap | Increase row cap or add aggregation |
