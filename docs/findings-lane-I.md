<!-- SPDX-License-Identifier: Apache-2.0 -->

# Lane I findings — Connection/Auth + Deploy

## CORS origins

### Dev (`deploy/envoy.yaml`)

The grpc-web CORS policy uses a safe_regex that permits exactly two localhost
origins:

```
^https?://localhost:(3000|8000)$
```

- `http://localhost:3000` — Vite dev server (`npm run dev`).
- `http://localhost:8000` — the static host listener in the compose stack.

The regex anchors at both ends so it cannot match `localhost:8000.evil.com`.
No wildcard `.*` is used.

### Prod (`deploy/envoy.prod.yaml`)

The grpc-web CORS policy uses an exact-string match:

```
exact: "https://YOUR-ADDIN-ORIGIN.example.com"
```

Replace the placeholder with the real add-in origin before deploying.
`scripts/render_envoy_prod.sh` can substitute `SCE_ADDIN_ORIGIN` from env.

The `authorization` header is included in `allow_headers` in both configs so a
Bearer token flows through CORS preflights to the Lua gate and then upstream.

## Token forwarding mechanism

1. **Task pane:** `saveToken(token)` writes to `OfficeRuntime.storage` (roaming,
   not in the .xlsx — DECISIONS #6). Falls back to an in-memory `Map` when
   `OfficeRuntime.storage` is unavailable (tests, offline). Never written to
   `Office.context.document.settings` or any cell.

2. **Envoy proxy:** the `authorization` header passes through to Spark Connect
   unchanged. In prod, a Lua filter rejects non-OPTIONS requests lacking
   `Authorization: Bearer <token>` (HTTP 401). OPTIONS preflights are let
   through unconditionally. The Lua gate checks presence only — replace with
   `jwt_authn` or `ext_authz` for signature validation (sketched in
   `deploy/envoy.prod.yaml`).

## COEP value: `credentialless` and why

Both dev and prod static-host listeners send:

```
Cross-Origin-Embedder-Policy: credentialless
```

`require-corp` would block Pyodide (jsDelivr CDN) and PyPI wheels because those
CDNs do not send `Cross-Origin-Resource-Policy` headers. `credentialless` grants
no-credentials cross-origin access without requiring CORP opt-in from the
resource server. `crossOriginIsolated` becomes `true` under `credentialless` in
Chromium-based Office hosts (WebView2, Edge) — the hosts we target. DECISIONS #2
locks this choice and CI asserts it.

## Connection string format

```
sc://host:port/;transport=grpcweb
```

`buildRemoteUri(cfg)` in `src/connection/connectionStore.ts` produces this.
TLS is a property of the Envoy proxy endpoint, not encoded in the `sc://` scheme.

| Scenario | URI |
|----------|-----|
| Local dev | `sc://localhost:8081/;transport=grpcweb` |
| Prod (TLS) | `sc://spark.example.com:8443/;transport=grpcweb` |

## Files created

| File | Purpose |
|------|---------|
| `src/connection/connectionStore.ts` | Config persistence, `buildRemoteUri`, token handling |
| `src/connection/connectionForm.ts` | DOM form render/validate helpers for Lane E |
| `deploy/compose.yaml` | Dev: Spark 4.0.0 + Envoy + static add-in host |
| `deploy/envoy.yaml` | Dev: grpc_web + localhost CORS + COEP credentialless |
| `deploy/compose.prod.yaml` | Prod overlay: TLS ports, certs mount, private Spark |
| `deploy/envoy.prod.yaml` | Prod: TLS + exact-origin CORS + Lua bearer gate |
| `deploy/README.md` | Ports, bring-up steps, CORS/COEP rationale, curl checks |
| `docs/findings-lane-I.md` | This file |
