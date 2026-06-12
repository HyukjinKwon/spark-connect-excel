<!-- SPDX-License-Identifier: Apache-2.0 -->

# Dependency Reuse - pyspark-connect-web

## What we reuse

This add-in is an **Excel host** around
[pyspark-connect-web](https://github.com/HyukjinKwon/pyspark-client-wasm)
(distribution name `pyspark-connect-web`, import name `pyspark_connect_web`).
We add Office.js glue; we do not fork Spark plumbing.

Two reuse mechanisms:

### 1. Same-origin runtime assets (Pyodide + wheels)

At runtime, `spark_excel_runtime.py` runs inside Pyodide and calls:

```python
import pyspark_connect_web as pcw
pcw.install()
```

The heavy runtime is served **same-origin** next to the app. A cross-origin CDN
does NOT work: under cross-origin isolation the worker's `importScripts()` of a
CDN Pyodide is blocked by COEP in Chromium, even with `credentialless` (this is
documented in pyspark-connect-web). These assets are version-matched to
pyspark-connect-web's own build and vendored into `public/` (git-ignored - they
are large):

As of release v0.1.0:

| Asset (served at site root) | What | Override global |
|-----------------------------|------|-----------------|
| `/pyodide/` (`pyodide.mjs`) | Pyodide distribution | `PCW_PYODIDE_INDEX_URL` |
| `/pyspark_client-4.1.2-py3-none-any.whl` | Spark Connect client wheel | `PCW_PYSPARK_WHEEL_URL` |
| `/pyspark_connect_web-0.1.0-py3-none-any.whl` | the pcw wheel (also on PyPI) | `PCW_WHEEL_URL` |

`micropip` still fetches the small pure deps (`protobuf`,
`googleapis-common-protos`, etc.) from PyPI at runtime. **Version constraint:**
`pyspark>=4.0,<4.2` (enforced by `pcw.install()`).

Run `npm run vendor:runtime` (see `scripts/vendor-runtime.mjs`) to fetch the
release glue + the pcw wheel into `public/`; it prints the remaining steps for
Pyodide and the `pyspark-client` wheel (version-matched to the release). Vite
copies `public/` into `dist/`. These assets are git-ignored
(`public/pyodide/`, `public/*.whl`). To host them elsewhere (same-origin),
override the globals above, e.g.:

```ts
await host.boot({ wheelUrl: "/cdn/pyspark_connect_web-0.1.0-py3-none-any.whl" });
```

### 2. Browser JS glue files (copied into `public/vendor/`)

Three small files must be served **same-origin** (they cannot come from a CDN):

| File | Upstream path | Role |
|------|---------------|------|
| `public/vendor/worker_bootstrap.js` | `pyspark_connect_web/worker/worker_bootstrap.js` | Web Worker entry: loads Pyodide, micropip-installs the wheel, allocates the SAB |
| `public/vendor/bridge.js` | `pyspark_connect_web/worker/bridge.js` | Main-thread `fetch` + SAB writeback for the Atomics/SAB handshake |
| `public/vendor/coi-serviceworker.js` | `pyspark_connect_web/jupyterlite/coi-serviceworker.js` | COI shim for header-less hosting (belt-and-suspenders) |

These files are copied verbatim from upstream and are **not edited**.
They carry a comment header in `public/vendor/README.md` recording their
upstream provenance.

---

## Pinned upstream version

The vendor files were copied from pyspark-connect-web at:

> **Upstream repository:** https://github.com/HyukjinKwon/pyspark-client-wasm  
> **Release:** `v0.1.0` - glue + pcw wheel taken from the
> `pyspark-connect-web-site-0.1.0.tgz` release asset (re-synced 2026-06-12).

Re-sync the vendor glue (and the same-origin asset versions) when a new release
ships.

---

## How to re-sync the vendor files

1. Clone or update the upstream repository:

```bash
git clone https://github.com/HyukjinKwon/pyspark-client-wasm /tmp/pcw
# or: cd /tmp/pcw && git pull
```

2. Copy the three files:

```bash
cp /tmp/pcw/pyspark_connect_web/worker/worker_bootstrap.js public/vendor/
cp /tmp/pcw/pyspark_connect_web/worker/bridge.js public/vendor/
cp /tmp/pcw/pyspark_connect_web/jupyterlite/coi-serviceworker.js public/vendor/
```

3. Update the pinned version line above.

4. Run `npm run build` and `npm run test:e2e` to verify nothing broke.

**Do not edit the vendor files.** If upstream has a bug, fix it upstream and
re-copy. This ensures we don't carry a private fork of core infrastructure.

---

## License

pyspark-connect-web is Apache-2.0 licensed. The copied files retain their
upstream Apache-2.0 license (the `SPDX-License-Identifier: Apache-2.0`
comment in each file covers this). This project is also Apache-2.0.

---

## Relationship to upstream

This project is an independent Excel add-in host that reuses pyspark-connect-web.
