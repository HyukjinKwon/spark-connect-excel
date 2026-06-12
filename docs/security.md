<!-- SPDX-License-Identifier: Apache-2.0 -->

# Security

## Threat model

spark-connect-excel is an Excel add-in that acts as a thin client to the user's
own Spark Connect cluster. It has **no backend server of its own**: all compute
runs on the user's cluster; the query client runs in the browser via Pyodide.
The security properties below follow from this architecture.

---

## Cross-origin isolation (COI)

### Why it matters

pyspark-connect-web's blocking `.collect()` relies on `SharedArrayBuffer` and
`Atomics.wait`, which the browser makes available only in **cross-origin-isolated**
contexts (`crossOriginIsolated === true`). Without this:

- `SharedArrayBuffer` is `undefined`.
- The Pyodide SAB bridge cannot function.
- No Spark queries can run.

This is not just a performance concern — COI is a hard requirement.

### How we achieve it (DECISIONS #1, #2)

Pyodide + pyspark-connect-web run inside an **Office Dialog window** opened via
`displayDialogAsync()`. That window is a real top-level browser context we serve
with:

```
Cross-Origin-Opener-Policy: same-origin
Cross-Origin-Embedder-Policy: credentialless
```

The `credentialless` value (DECISIONS #2) is deliberate: `require-corp` would
block Pyodide (jsDelivr CDN) and the pyspark-connect-web wheel (PyPI CDN) because
those origins do not send `Cross-Origin-Resource-Policy` headers.
`credentialless` grants cross-origin access to resources that carry no credentials
(cookies/auth), covering CDN wheels without weakening isolation for authenticated
resources.

CI asserts both headers are present in `vite.config.ts` and the `deploy/` stack.
The automated Playwright test (`tests/e2e/coi.spec.ts`) asserts
`crossOriginIsolated === true` end-to-end against a real built page.

---

## Bearer token handling (DECISIONS #6)

**Invariant: bearer tokens NEVER touch a cell or `Office.context.document.settings`.**

### Where tokens are stored

| Location | API | Visibility |
|----------|-----|------------|
| `OfficeRuntime.storage` | `OfficeRuntime.storage.setItem(key, value)` | Roaming, per-user, **not** embedded in the .xlsx file |
| In-memory fallback | JavaScript `Map` | Session-scoped; wiped on page reload. Used when `OfficeRuntime` is unavailable (tests, offline) |

`Office.context.document.settings` stores only the **non-secret** connection
config (host, port, TLS flag) so the workbook can pre-populate the form on
reopen. The token is never part of that settings value.

### Why this matters

An `.xlsx` file is a zip archive. `Office.context.document.settings` values are
embedded in the archive (`xl/customXml/` or `docProps/`). A workbook saved to a
file share, emailed to a colleague, or committed to version control would expose
any token stored there. `OfficeRuntime.storage` is a platform-level secret store
opaque to the file.

### Token flow to Spark

The token is appended to the `sc://` URI as a `token=` parameter before the
Spark session is created:

```
sc://host:port/;transport=grpcweb;token=<bearer-value>
```

PySpark's `DefaultChannelBuilder` reads the `token=` param and maps it to the
`authorization: Bearer <token>` grpc channel metadata, which the pyspark-connect-web
patch forwards as a grpc-web HTTP header to the Envoy proxy.

The Envoy proxy is the enforcement point:
- In **dev**: no auth gate (localhost only).
- In **prod**: a Lua filter rejects non-OPTIONS requests lacking
  `Authorization: Bearer <value>` with HTTP 401. Replace with `jwt_authn` or
  `ext_authz` for signature validation.

---

## CORS

### Dev stack (`deploy/envoy.yaml`)

```
allow-origin regex: ^https?://localhost:(3000|8000)$
```

Anchored at both ends — `localhost:8000.evil.com` does not match.

### Prod stack (`deploy/envoy.prod.yaml`)

```
allow-origin exact: https://YOUR-ADDIN-ORIGIN.example.com
```

No wildcard; exact-string match only. The `authorization` header is in
`allow_headers` so Bearer tokens flow through CORS preflights.

---

## TLS

For production deployments:

- The Envoy proxy terminates TLS on port 8443.
- The `sc://` URI still uses `transport=grpcweb` — TLS is implied by the
  proxy, not the URI scheme.
- The browser requires HTTPS for `crossOriginIsolated` off localhost.
- Provide a real TLS certificate (`deploy/certs/tls.crt` + `tls.key`);
  self-signed certs are acceptable for staging only.

---

## Untrusted-server caveat

The add-in connects to whatever Spark Connect endpoint the user enters. It does
not validate server certificates beyond what the browser enforces. In production:

- Use a CA-signed certificate at the Envoy proxy.
- The user's Excel host enforces standard HTTPS certificate validation.
- Do not disable certificate verification.

**The add-in runs the user's own queries against the user's own cluster.** It
does not send queries to any Anthropic, Microsoft, or third-party backend. The
query text, connection URI, and results are only ever transmitted between the
user's browser and the user's Spark Connect endpoint.

---

## Source integrity

The Python runtime (`spark_excel_runtime.py`) is bundled at build time via
Vite's `?raw` import and loaded into Pyodide via a `base64`-encoded `exec()`.
The source is part of the open-source repository; there is no external Python
package fetch at runtime for the runtime module itself (only for
`pyspark-connect-web` which is fetched from PyPI via micropip).

The three vendor JS files (`public/vendor/`) are copied verbatim from the
pyspark-connect-web repository (Apache-2.0). They are served same-origin and
are not fetched from a CDN. See `docs/reuse.md`.

---

## Responsible disclosure

This is an open-source community project. Please report security issues via
GitHub Issues or email the repository owner. There is no formal bug bounty
program.
