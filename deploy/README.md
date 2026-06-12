<!-- SPDX-License-Identifier: Apache-2.0 -->

# deploy/ - Spark Connect + Envoy grpc-web proxy for spark-connect-excel

This directory brings up the **server side** of the spark-connect-excel add-in:
a Spark 4.x Connect server fronted by an Envoy `grpc_web` proxy, plus a static
host that serves the built Excel add-in (`dist/`) with the mandatory
cross-origin-isolation headers.

```
Excel add-in dialog (Office Dialog window, COEP credentialless)
   |  sc://localhost:8081/;transport=grpcweb   (grpc-web over fetch)
   v
Envoy :8081  --grpc_web filter-->  Spark Connect :15002 (gRPC/HTTP2)
Envoy :8000  -- static add-in (dist/)  with COOP/COEP -->  halverneus static server
```

## Ports

| Port | Service | What |
|------|---------|------|
| 8081 | Envoy (dev) | grpc-web endpoint. Add-in URI: `sc://localhost:8081/;transport=grpcweb` |
| 8000 | Envoy (dev) | Add-in static host (`dist/`). `Cross-Origin-Opener-Policy: same-origin` + `Cross-Origin-Embedder-Policy: credentialless`. Point manifest `SourceLocation` here for local demo. |
| 15002 | spark-connect | Raw gRPC (published in dev only; private in prod). |
| 9901 | Envoy admin | Dev only; NOT published in prod. |
| 8443 | Envoy (prod) | grpc-web over TLS. Prod add-in URI: `sc://YOUR-HOST:8443/;transport=grpcweb` |
| 8444 | Envoy (prod) | Add-in static host over TLS with COOP/COEP. |
| 8089 | Envoy (prod) | Health/readiness probe (`/healthz`, `/ready`). |

## Bring it up (dev)

```bash
# Build the add-in first so dist/ is populated:
npm run build          # or: vite build

docker compose -f deploy/compose.yaml up
# wait ~60s for Spark Connect to become healthy (gRPC port open)
```

Point the add-in at the grpc-web proxy:

```
sc://localhost:8081/;transport=grpcweb
```

In the connection form: host = `localhost`, port = `8081`, TLS = off.

## Bring it up (prod hardening overlay)

```bash
# 1. Provide a TLS cert:
mkdir -p deploy/certs
# Production: use a real cert (Let's Encrypt / your CA):
#   deploy/certs/tls.crt   (full chain)
#   deploy/certs/tls.key   (private key, chmod 600)
# Quick self-signed for staging only:
PCW_PUBLIC_HOST=spark.example.com scripts/gen_dev_cert.sh

# 2. Set the add-in origin in envoy.prod.yaml (or render from env):
SCE_PUBLIC_HOST=spark.example.com \
SCE_ADDIN_ORIGIN=https://addin.example.com \
scripts/render_envoy_prod.sh > deploy/envoy.prod.rendered.yaml

# 3. Bring up:
docker compose -f deploy/compose.yaml -f deploy/compose.prod.yaml up -d
```

## How the add-in connects

The add-in uses the `sc://host:port/;transport=grpcweb` URI scheme understood
by pyspark-connect-web. TLS is determined by the scheme used to reach the Envoy
proxy (plaintext port 8081 in dev; HTTPS port 8443 in prod). The `sc://` form
always carries `transport=grpcweb` regardless of TLS - the proxy negotiates TLS
externally.

Example URIs:

| Scenario | URI |
|----------|-----|
| Local dev | `sc://localhost:8081/;transport=grpcweb` |
| Prod (TLS) | `sc://spark.example.com:8443/;transport=grpcweb` |

## Pointing the manifest at the static host (non-localhost demo)

In `manifest.xml` change `SourceLocation` to the prod static host:

```xml
<SourceLocation DefaultValue="https://addin.example.com/taskpane.html"/>
```

The static host at `:8444` (prod) or `:8000` (dev) serves the full `dist/`
directory. It must be reachable from the user's browser (Excel on Web / Windows
WebView2 / Mac WKWebView) over HTTPS for prod (or HTTP for localhost dev).

## COOP/COEP and why `credentialless`

The Office Dialog window (which hosts Pyodide + pyspark-connect-web) requires
`crossOriginIsolated === true` to use `SharedArrayBuffer` - the backbone of
pyspark-connect-web's blocking `.collect()`. That requires:

```
Cross-Origin-Opener-Policy: same-origin
Cross-Origin-Embedder-Policy: credentialless   <- NOT require-corp
```

We use `credentialless` (DECISIONS #2) rather than `require-corp` because:

- Pyodide loads from **jsDelivr CDN** and Python packages from **PyPI CDN**.
  Neither of those origins sends a `Cross-Origin-Resource-Policy` header.
- `require-corp` would block any cross-origin resource that does not opt in
  with CORP; that includes Pyodide itself.
- `credentialless` grants cross-origin access to resources that do not carry
  credentials (cookies/auth), which covers the CDN wheels.
- Chromium-based Office hosts (WebView2, Edge, Chrome) fully support
  `credentialless`; it is the correct value for our use-case.

## CORS origins

| Config | CORS origin(s) allowed |
|--------|------------------------|
| Dev (`envoy.yaml`) | `http://localhost:3000` (Vite dev server), `http://localhost:8000` (static host) - via regex `^https?://localhost:(3000\|8000)$` |
| Prod (`envoy.prod.yaml`) | `https://YOUR-ADDIN-ORIGIN.example.com` - exact-string match only, no wildcard |

The `authorization` request header is included in `allow_headers` on both
configs so the Bearer token flows through CORS preflights to the Envoy proxy,
which forwards it upstream to Spark Connect (DECISIONS #6 enforcement point).

## Verifying CORS / grpc-web / COOP+COEP without a browser

```bash
# 1. Verify CORS preflight on the grpc-web port:
curl -i -X OPTIONS http://localhost:8081/spark.connect.SparkConnectService/ExecutePlan \
  -H "Origin: http://localhost:8000" \
  -H "Access-Control-Request-Method: POST" \
  -H "Access-Control-Request-Headers: x-grpc-web,content-type,authorization"
# Expect: access-control-allow-origin, access-control-allow-headers in the response.

# 2. Verify CORS preflight with the Vite dev server origin:
curl -i -X OPTIONS http://localhost:8081/spark.connect.SparkConnectService/ExecutePlan \
  -H "Origin: http://localhost:3000" \
  -H "Access-Control-Request-Method: POST" \
  -H "Access-Control-Request-Headers: x-grpc-web,content-type,authorization"
# Expect: access-control-allow-origin: http://localhost:3000

# 3. Verify isolation headers on the static add-in host:
curl -sI http://localhost:8000/ | grep -i 'cross-origin'
# Expect:
#   Cross-Origin-Opener-Policy: same-origin
#   Cross-Origin-Embedder-Policy: credentialless
#   Cross-Origin-Resource-Policy: cross-origin

# 4. Verify the bearer-token gate in prod (should reject):
curl -i -X POST https://localhost:8443/spark.connect.SparkConnectService/ExecutePlan \
  --insecure \
  -H "Origin: https://addin.example.com" \
  -H "x-grpc-web: 1" \
  -H "content-type: application/grpc-web+proto"
# Expect: HTTP 401

# 5. Verify the bearer-token gate passes with a token:
curl -i -X POST https://localhost:8443/spark.connect.SparkConnectService/ExecutePlan \
  --insecure \
  -H "Origin: https://addin.example.com" \
  -H "x-grpc-web: 1" \
  -H "content-type: application/grpc-web+proto" \
  -H "Authorization: Bearer mytoken"
# Expect: HTTP 200 (or gRPC error - not 401)
```

## Token forwarding mechanism

The `authorization` header is:
1. Listed in `allow_headers` in both dev and prod CORS configs so the browser
   CORS preflight allows it.
2. NOT stripped by any Envoy filter - it passes through to the Spark Connect
   upstream unchanged.
3. Enforced (presence-only) in prod by a Lua filter before the router, which
   returns 401 if the header is absent or does not start with `Bearer `.
4. Stored **only** in `OfficeRuntime.storage` (roaming, not in the .xlsx) by
   the task-pane's `connectionStore.saveToken()` - never in document settings
   or a cell (DECISIONS #6).

Replace the Lua gate with `jwt_authn` or `ext_authz` for signature validation.
See the commented options at the bottom of `envoy.prod.yaml`.

## Image / version pins

| Component | Pin | Notes |
|-----------|-----|-------|
| Spark Connect server | `apache/spark:4.0.0` | Bundles Connect; matches `pyspark>=4.0` |
| Spark Connect package | `org.apache.spark:spark-connect_2.13:4.0.0` | Must match Spark + Scala version |
| Envoy | `envoyproxy/envoy:v1.31-latest` | Has `envoy.filters.http.grpc_web` |
| Static host | `halverneus/static-file-server:v1.8.10` | Serves `dist/` on :80 |

## Dev vs prod summary

| Concern | Dev | Prod |
|---------|-----|------|
| Transport | plaintext HTTP | TLS (HTTPS), TLSv1_2+ |
| Ports | 8081 grpc-web, 8000 static, 9901 admin | 8443 grpc-web, 8444 static, 8089 probe |
| CORS | localhost:3000 + localhost:8000 (regex) | Exact `https://YOUR-ADDIN-ORIGIN` only |
| Auth | none | Bearer presence gate (Lua); replace with jwt_authn/ext_authz |
| Spark gRPC | published `15002:15002` | NOT published (private) |
| COEP | `credentialless` | `credentialless` (same; see rationale above) |
