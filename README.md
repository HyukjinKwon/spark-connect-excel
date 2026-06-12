<!-- SPDX-License-Identifier: Apache-2.0 -->

# Spark Connect for Excel

[![CI](https://github.com/HyukjinKwon/spark-connect-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/HyukjinKwon/spark-connect-excel/actions/workflows/ci.yml)
[![E2E (COI gate)](https://github.com/HyukjinKwon/spark-connect-excel/actions/workflows/e2e.yml/badge.svg)](https://github.com/HyukjinKwon/spark-connect-excel/actions/workflows/e2e.yml)
[![License](https://img.shields.io/badge/license-Apache--2.0-green.svg)](LICENSE)
[![Docs](https://img.shields.io/badge/docs-GitHub%20Pages-blue)](https://hyukjinkwon.github.io/spark-connect-excel/)

**Power Query, but the engine is your Spark cluster.**

An Excel add-in that runs a Spark SQL query against your own
[Spark Connect](https://spark.apache.org/docs/latest/spark-connect-overview.html)
cluster, lands the result directly in a worksheet range, refreshes it, and
charts it - with **no backend server of its own**. The query client runs
entirely in-browser via
[pyspark-connect-web](https://github.com/HyukjinKwon/pyspark-client-wasm)
(real PySpark, in Pyodide). See the
[full documentation](https://hyukjinkwon.github.io/spark-connect-excel/).

<p align="center">
  <img src="docs/screenshots/web-demo-live.png" alt="The web demo connected to a live Spark Connect server, running a Spark SQL query and rendering the result as a table and a bar chart." width="820" />
  <br/>
  <em>A live Spark SQL query in the browser - real result table and chart, connected to a Spark Connect server. Captured by the end-to-end CI run.</em>
</p>

## Features

- **SQL in Excel** - write Spark SQL in the task pane; results land as typed cells.
- **Typed ranges** - Spark's schema drives Excel number formats (dates, decimals, integers).
- **Truncation guard** - row cap (default 10k) with a visual banner when the result is clipped.
- **Refresh** - rebind the query to its range; one click updates stale data.
- **Native charts** - auto-inferred chart type (line for time-series, column for categories, scatter for numeric pairs).
- **No backend** - compute runs on your Spark cluster; the add-in is static HTML/JS.
- **Secure token handling** - bearer tokens never touch a cell or the workbook file.
- **Zero-install web demo** - a standalone page (`/demo`) that runs **SQL _or_ PySpark**
  in the browser, no Excel required - the easiest way to try it.

## Quickstart

### Prerequisites

- Node 20, plus `python3` + `pip`, `git`, `curl`, `tar` (the quickstart vendors
  Pyodide + the wheels the in-browser engine needs - it does this for you).
- Excel **desktop** (Microsoft 365 / 2019, Windows or Mac) to run the engine. Excel on the web loads the task pane but cannot run the engine (see Compatibility). The standalone web demo runs in any Chromium browser.
- A running Spark Connect server (see below).

### Start a Spark Connect server

The demo and the add-in both talk to a Spark Connect server through an Envoy
grpc-web proxy. Start one first, either way:

```bash
# Docker - Spark Connect + Envoy (grpc-web at :8081) in one command:
docker compose -f deploy/compose.yaml up
```

No Docker? Run a server locally with PySpark (Java 17) and put Envoy in front of
it - see the
[installation guide](https://hyukjinkwon.github.io/spark-connect-excel/installation/).

> The quickstart below vendors the engine assets (Pyodide + wheels) and serves
> the add-in over HTTPS. This exact stack - Pyodide + grpc-web + Spark Connect -
> is exercised end to end by the `e2e-full` CI gate on every push, which runs a
> live Spark SQL query and renders the chart shown above.

### Try it without Excel - the web demo

A standalone page, no Excel and no sideload. One command clones, installs, and
serves the add-in:

```bash
curl -fsSL https://raw.githubusercontent.com/HyukjinKwon/spark-connect-excel/main/scripts/quickstart.sh | bash
```

Open `https://localhost:3000/demo/demo.html`, point it at your Spark Connect
server, and run a query. See the
[usage guide](https://hyukjinkwon.github.io/spark-connect-excel/usage/).

### Try with Excel on Windows / Mac desktop (sideloads) - runs the engine

Excel **desktop** is the host that runs the engine. With a Spark Connect server
running:

```bash
curl -fsSL https://raw.githubusercontent.com/HyukjinKwon/spark-connect-excel/main/scripts/quickstart.sh | bash -s -- desktop
```

This serves the add-in and opens Excel with it sideloaded. The **Apache Spark**
button (group **Spark Connect**) appears on the Home ribbon.

### Excel on the web - task pane only (engine does not run)

You can sideload the add-in on Excel on the web (**Home -> Add-ins -> MY ADD-INS
-> "Choose your add-in manifest"** -> `manifest.xml`) and the task pane loads,
but **the engine will not start there**: it needs `SharedArrayBuffer`, which
requires a cross-origin-isolated dialog window that Excel on the web does not
permit (you get a "security zone" error). Use Excel **desktop** for the engine,
or run the standalone **web demo** above, which works in any Chromium browser.
(This is the same in Chrome and Edge - it is a platform limit, not a browser
setting.) Full guides are on the
[documentation site](https://hyukjinkwon.github.io/spark-connect-excel/).

## Compatibility

| Component | Supported |
|-----------|-----------|
| Excel (engine) | Microsoft 365 / 2019 desktop - Windows or Mac |
| Excel on the web | task pane UI only; engine needs desktop (no cross-origin-isolated dialog) |
| Excel API requirement | ExcelApi 1.12, DialogApi 1.2 |
| Spark | 4.x (Spark Connect; `apache/spark:4.0.0` in the deploy stack) |
| PySpark | `>=4.0,<4.2` (enforced by `pcw.install()`) |
| Browser engine | Chromium-based (WebView2, Edge, Chrome) - `COEP: credentialless` is Chromium-only |
| Node | 20 LTS |
| Python | 3.11+ (for local dev/tests; Pyodide 0.28+ in the browser) |
