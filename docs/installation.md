<!-- SPDX-License-Identifier: Apache-2.0 -->

# Installation & Development Guide

## Prerequisites

| Requirement | Version | Notes |
|-------------|---------|-------|
| Node.js | 20 LTS | `node --version` |
| npm | 10+ | Comes with Node 20 |
| Python | 3.11+ | For running the Python unit tests locally |
| Docker (optional) | 20+ | Only needed for the full stack (Spark Connect + Envoy + static host); a local server just needs Java 17 + `pip install "pyspark[connect]"` |
| Excel (to run the engine) | 2019 / Microsoft 365 | Windows or Mac **desktop** (Excel on the web loads the task pane but cannot run the engine - see below) |

> **Browser requirement:** The add-in requires a Chromium-based Excel host
> (Windows WebView2, Edge, Chrome). `COEP: credentialless` - the mode needed
> for `SharedArrayBuffer` - is a Chromium feature. Safari and Firefox are not
> supported in v0.

---

## Install dependencies

```bash
git clone https://github.com/HyukjinKwon/spark-connect-excel
cd spark-connect-excel
npm install
```

---

## Development server

### Plain HTTP (for unit tests and the COI gate)

```bash
npm run dev
# Vite dev server starts on http://localhost:3000
# Serves both pages:
#   http://localhost:3000/taskpane/taskpane.html
#   http://localhost:3000/dialog/dialog.html
```

The dev server sets the mandatory COI headers:
```
Cross-Origin-Opener-Policy: same-origin
Cross-Origin-Embedder-Policy: credentialless
```

### HTTPS (required for Excel on Windows/Mac)

Office add-ins on Windows (WebView2) and Mac (WKWebView) require HTTPS for the
task pane. Use `office-addin-dev-certs` to generate a local certificate:

```bash
npx office-addin-dev-certs install   # once - installs an OS-trusted local cert
npm run dev:https                    # serves https://localhost:3000 with COI headers
```

`npm run dev:https` sets `HTTPS=true`, which makes Vite load the trusted cert
from `office-addin-dev-certs`. The cert is trusted by the OS so WebView2/WKWebView
accept it. (Plain `npm run dev` stays HTTP - used for unit tests / the COI gate.)

---

## Build

```bash
npm run build
# Produces dist/ - taskpane/taskpane.html + dialog/dialog.html + assets
```

Run the full check suite before building:

```bash
npm run typecheck   # TypeScript strict-mode check
npm run lint        # ESLint + Prettier
npm test            # Vitest unit tests
python -m pytest python/tests -q   # Python unit tests
npm run build       # tsc --noEmit + vite build
```

---

## Sideload into Excel

Sideloading installs the add-in directly, without an add-in store.

### One command (quickstart)

The fastest path - clones, installs, sets up the local HTTPS cert, and serves
the add-in (then upload the manifest as below):

```bash
# Serve the add-in (then sideload in Excel desktop, below):
curl -fsSL https://raw.githubusercontent.com/HyukjinKwon/spark-connect-excel/main/scripts/quickstart.sh | bash

# Windows / Mac desktop (also sideloads + opens Excel):
curl -fsSL https://raw.githubusercontent.com/HyukjinKwon/spark-connect-excel/main/scripts/quickstart.sh | bash -s -- desktop
```

The script is `scripts/quickstart.sh`. Prefer to run the steps yourself? Follow
the manual instructions below.

!!! warning "The engine runs on Excel desktop, not Excel on the web"
    The engine runs in a **cross-origin-isolated dialog window** (it needs
    `SharedArrayBuffer`). Excel **desktop** (Windows/Mac) hosts that dialog in an
    embedded webview and works. **Excel on the web cannot** - it refuses to open
    a cross-origin-isolated dialog ("the dialog's domain and the add-in host's
    domain are not in the same security zone"), the same in Chrome and Edge. On
    Excel on the web the task pane loads but the engine will not start. To run
    in-browser without Excel, use the standalone **web demo**
    (`/demo/demo.html`), which is a top-level page and works in any Chromium
    browser.

### Excel on the web - task pane only (engine does not run)

You can still sideload to inspect the UI:

1. Run `npm run dev:https` (and `npx office-addin-dev-certs install` once).
2. Open **Excel on the web** in Microsoft Edge or Google Chrome.
3. **Home -> Add-ins** (the Add-ins button is on the **Home** tab in current
   Excel on the web; older builds put it under **Insert**). In the Office Add-ins
   dialog, open the **MY ADD-INS** tab and click **"Choose your add-in
   manifest"** (a.k.a. "Upload My Add-in"), then pick `manifest.xml`.
4. The **Apache Spark** button appears on the Home ribbon, but opening the engine
   window will fail with the "security zone" error above. Use Excel desktop.

**To let other people install it**, the add-in must be hosted where they can
reach it - host the built bundle (`npm run build` -> `dist/`) on an HTTPS origin
that sends the COI headers, then point the manifest at it:

```bash
npm run build:manifest -- --origin https://your-addin-host.example.com
# share dist/manifest.xml - recipients do the same 3 upload steps
```

(For a quick demo from your own machine, expose `https://localhost:3000` with a
tunnel like `ngrok http 3000` and `build:manifest` against the tunnel URL.)

### Windows desktop (runs the engine)

```bash
npx office-addin-debugging start manifest.xml
# Installs the dev cert, starts the dev server, opens Excel with the add-in
```

Or manually: **File > Options > Trust Center > Trusted Add-in Catalogs**, add a
shared-folder/URL catalog, then **Home (or Insert) > Add-ins > My Add-ins >
Shared Folder**.

### Mac desktop (runs the engine)

```bash
npx office-addin-debugging start manifest.xml
# office-addin-debugging copies the manifest to the correct location
```

Or manually: copy `manifest.xml` to
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` (create the
`wef/` directory if it doesn't exist). Restart Excel.

> The engine runs on Excel **desktop** (Windows or Mac). The Office.js error's
> suggestion to "configure your browser" / "try a different browser" refers to
> legacy Internet Explorer **Security Zones** (a Windows-only IE concept) and
> does **not** apply to Chrome/Edge on macOS - there is no such setting to change,
> and it would not fix this anyway. The real cause is that the engine's dialog
> must be cross-origin-isolated, which Excel on the web does not allow. Use Excel
> desktop, or the standalone web demo for in-browser use.

---

## Point the add-in at the Spark Connect stack

The connection form in the task pane accepts:

| Field | Dev default | Prod example |
|-------|-------------|--------------|
| Host | `localhost` | `spark.example.com` |
| Port | `8081` | `8443` |
| TLS | off | on |
| Token | (empty for dev) | Bearer token for prod |

### A Spark Connect server - two ways

**Local (no Docker) - easiest for dev/tests.** Install PySpark with the
Connect extra (needs Java 17) and start a local server:

```bash
pip install "pyspark[connect]==4.0.0"
# start a Spark Connect server on localhost:15002
start-connect-server.sh --packages org.apache.spark:spark-connect_2.13:4.0.0
```

This is what the Python integration tests use (`SparkSession.builder.remote("local[*]")`). For the browser add-in you also need the Envoy grpc-web proxy in front of it (the deploy stack provides one, or run a local Envoy).

**Full stack (Docker) - one command.** Brings up Spark Connect + Envoy + a static host together:

```bash
docker compose -f deploy/compose.yaml up
# Wait ~60s for Spark Connect to report healthy
```

Then in the task pane: host = `localhost`, port = `8081`, TLS = off. No token needed
for the dev stack.

Connection URI produced by the add-in:
```
sc://localhost:8081/;transport=grpcweb
```

---

## Available npm scripts

| Script | What it does |
|--------|-------------|
| `npm run dev` | Start Vite dev server |
| `npm run build` | Type-check + Vite build -> `dist/` |
| `npm run preview` | Serve `dist/` on port 3000 (for e2e tests) |
| `npm run typecheck` | `tsc --noEmit` only |
| `npm run lint` | ESLint + Prettier check |
| `npm run lint:fix` | ESLint + Prettier auto-fix |
| `npm test` | Vitest unit tests |
| `npm run test:watch` | Vitest in watch mode |
| `npm run test:py` | Python pytest |
| `npm run test:e2e` | Playwright e2e tests |
| `npm run validate` | `office-addin-manifest validate manifest.xml` |

---

## Running the COI gate (e2e)

```bash
npm run build
npx playwright install --with-deps chromium
npx playwright test coi.spec.ts --config tests/e2e/playwright.config.ts
```

This is the automated infrastructure check described in `docs/architecture.md`
and `tests/e2e/README.md`.
