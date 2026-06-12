<!-- SPDX-License-Identifier: Apache-2.0 -->

# Spark Connect for Excel

**Power Query, but the engine is your Spark cluster.**

Run a Spark SQL query (or PySpark) against your own
[Spark Connect](https://spark.apache.org/docs/latest/spark-connect-overview.html)
cluster, land the result in a worksheet range, refresh it, and chart it - with
**no backend server**. The query client runs entirely in-browser via
[pyspark-connect-web](https://github.com/HyukjinKwon/pyspark-client-wasm)
(real PySpark, in Pyodide).

<p align="center">
  <img src="demo.svg" alt="Type Spark SQL, run it, rows land in the grid with a chart - no backend." width="820" />
</p>

## Highlights

- **SQL in Excel** - results land as typed cells; Spark's schema drives number formats.
- **Refresh & charts** - rebind a query to its range; one-click native charts.
- **No backend** - compute runs on your cluster; the add-in is static HTML/JS.
- **Zero-install web demo** - a standalone page with a SQL / Python toggle.

## Get started

- [Installation & sideload](installation.md) - load it into Excel (web or desktop).
- [Usage](usage.md) - connect, query, refresh, chart.
- [Distribution](distribution.md) - host it / share it / AppSource.

## Under the hood

- [Architecture](architecture.md) - task pane -> COI dialog -> Pyodide -> Envoy -> Spark.
- [Security](security.md) - cross-origin isolation, token handling, CORS/TLS.
- [Dependency reuse](reuse.md) - what we vendor from pyspark-connect-web.

---

Apache-2.0. Source: <https://github.com/HyukjinKwon/spark-connect-excel>.
