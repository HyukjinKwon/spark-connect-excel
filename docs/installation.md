<!-- SPDX-License-Identifier: Apache-2.0 -->

# Installation & Development Guide

## Prerequisites

| Requirement | Version | Notes |
|-------------|---------|-------|
| Node.js | 20 LTS | `node --version` |
| npm | 10+ | Comes with Node 20 |
| Python | 3.11+ | For running the Python unit tests locally |
| Docker (optional) | 20+ | For the local Spark Connect + Envoy stack |
| Excel | 2019 / Microsoft 365 | Windows, Mac, or Excel on the web |

> **Browser requirement:** The add-in requires a Chromium-based Excel host
> (Windows WebView2, Edge, Chrome). `COEP: credentialless` — the mode needed
> for `SharedArrayBuffer` — is a Chromium feature. Safari and Firefox are not
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
npx office-addin-dev-certs install
npm run dev
# or if a dev:https script is added to package.json:
# HTTPS=true npm run dev
```

The dev cert is trusted by the OS so WebView2/WKWebView accept it.

---

## Build

```bash
npm run build
# Produces dist/ — taskpane/taskpane.html + dialog/dialog.html + assets
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

Sideloading lets you test the add-in without publishing it to AppSource.
See `scripts/sideload.md` for the quick-start command, or follow the steps below.

### Windows

```bash
npx office-addin-debugging start manifest.xml
# Opens Excel with the add-in sideloaded and the dev server running
```

Or manually: copy `manifest.xml` to `%APPDATA%\Microsoft\Excel\XLSTART\` (Excel
must be restarted). Alternatively, use the **Insert > Add-ins > My Add-ins >
Upload My Add-in** dialog.

### Mac

```bash
npx office-addin-debugging start manifest.xml
# On Mac, office-addin-debugging copies the manifest to the correct location
```

Or manually: copy `manifest.xml` to
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` (create the
`wef/` directory if it doesn't exist). Restart Excel.

### Excel on the web (Microsoft 365)

1. Open Excel on the web.
2. Go to **Insert > Add-ins > Upload My Add-in**.
3. Upload `manifest.xml`.

The add-in must be hosted at the `SourceLocation` in `manifest.xml` (default:
`https://localhost:3000`). For Excel on the web, the add-in host must be
accessible from the internet or via a tunnel (e.g. `ngrok`).

---

## Point the add-in at the Spark Connect stack

The connection form in the task pane accepts:

| Field | Dev default | Prod example |
|-------|-------------|--------------|
| Host | `localhost` | `spark.example.com` |
| Port | `8081` | `8443` |
| TLS | off | on |
| Token | (empty for dev) | Bearer token for prod |

To bring up the local Spark Connect + Envoy stack:

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
| `npm run build` | Type-check + Vite build → `dist/` |
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
