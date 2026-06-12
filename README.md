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

- Node 20.
- Java 17 + Python 3.11 (to run a local Spark Connect server) - OR Docker (for the full Spark + Envoy stack).
- A Chromium-based Excel host: Excel on Windows / Microsoft 365, or Excel on the web in Edge or Chrome.

### Install and dev

```bash
git clone https://github.com/HyukjinKwon/spark-connect-excel
cd spark-connect-excel
npm install
npx office-addin-dev-certs install     # once - OS-trusted local HTTPS cert
npm run dev:https                      # serves https://localhost:3000 (COI headers)
```

### Bring up the Spark Connect stack

```bash
docker compose -f deploy/compose.yaml up
# Wait ~60s for Spark Connect to become healthy
```

### Try it without Excel - the web demo

The lowest-friction way to try Spark-in-the-browser: open
`https://localhost:3000/demo/demo.html` (with `npm run dev:https` running). It's
a standalone page - **no Excel, no sideload** - with a connection form and a
**SQL / Python** toggle:

- **SQL mode** -> results render as a typed table.
- **Python mode** -> run real PySpark (`spark.range(10).filter("id % 2 = 0").toPandas()`);
  `spark` is the connected session.

To host it for others, build (`npm run build`) and deploy `dist/` to a static host
that sets COI headers. `dist/_headers` (Netlify / Cloudflare Pages) and
`dist/staticwebapp.config.json` (Azure Static Web Apps) are included and set them
for you.

### Install into Excel (sideload)

First get the project and start the local add-in server (one time):

```bash
git clone https://github.com/HyukjinKwon/spark-connect-excel
cd spark-connect-excel
npm install
npx office-addin-dev-certs install   # trusted local HTTPS cert (once)
```

**Excel on the web - recommended (no admin, fewest clicks):**

```bash
npm run dev:https                    # serve the add-in at https://localhost:3000
```

Then, in **Excel on the web** (Edge or Chrome): **Insert -> Add-ins ->
Upload My Add-in -> choose `manifest.xml`**. The **Spark SQL** button appears on
the Home ribbon.

> Terminal-only web sideload: a CLI path exists
> (`npx office-addin-debugging start manifest.xml web --document <workbook-url>`)
> but it needs a workbook URL on OneDrive/SharePoint and an M365 sign-in, so it
> is not actually simpler than the upload above. For local dev, the manual
> upload is the fastest.

**Windows / Mac desktop (one command):**

```bash
git clone https://github.com/HyukjinKwon/spark-connect-excel
cd spark-connect-excel
npm install
npx office-addin-dev-certs install                       # trusted cert (once)
npm run dev:https &                                       # serve the add-in
npx office-addin-debugging start manifest.xml desktop --app excel   # sideload + open Excel
```

`office-addin-debugging start` copies the manifest to the right place and opens
Excel with the add-in loaded; `office-addin-debugging stop manifest.xml` removes
it. (Manual paths are in [scripts/sideload.md](scripts/sideload.md).)

> Requires a Chromium-based Excel host (Excel on Windows / Microsoft 365, or
> Excel on the web in Edge or Chrome). The add-in shows a clear message on
> unsupported hosts.

To let *other people* install it without cloning, host the built bundle
(`npm run build` -> `dist/`) on an HTTPS origin with COI headers, regenerate the
manifest (`npm run build:manifest -- --origin https://your-host`), and share that
`manifest.xml`. See [docs/distribution.md](docs/distribution.md).

Full installation guide: [docs/installation.md](docs/installation.md)

## Compatibility

| Component | Supported |
|-----------|-----------|
| Excel | 2019 / Microsoft 365 - Windows, Mac, Excel on the web |
| Excel API requirement | ExcelApi 1.12, DialogApi 1.2 |
| Spark | 4.x (Spark Connect; `apache/spark:4.0.0` in the deploy stack) |
| PySpark | `>=4.0,<4.2` (enforced by `pcw.install()`) |
| Browser engine | Chromium-based (WebView2, Edge, Chrome) - `COEP: credentialless` is Chromium-only |
| Node | 20 LTS |
| Python | 3.11+ (for local dev/tests; Pyodide 0.28+ in the browser) |
