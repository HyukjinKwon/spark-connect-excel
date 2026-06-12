<!-- SPDX-License-Identifier: Apache-2.0 -->

# Distribution

## Two distribution paths

### 1. Sideload (developer / self-hosted / private team)

Sideloading installs the add-in directly from a manifest, no add-in store
involved. It is the path for:

- Development and testing
- Private team deployments
- IT-managed enterprise rollouts via Centralized Deployment (below)

See [installation.md](installation.md) for step-by-step instructions.

### 2. AppSource (public distribution)

AppSource submission makes the add-in available to any Microsoft 365 user via
**Insert > Get Add-ins**. This path requires a Microsoft Partner Center account,
the bundle hosted on a public HTTPS origin with the COI headers (below), and
passing Microsoft's automated + manual validation checklist.

---

## Hosting requirements

However it is hosted, the add-in bundle (`dist/`) **must**
be hosted with the following HTTP response headers:

```
Cross-Origin-Opener-Policy: same-origin
Cross-Origin-Embedder-Policy: credentialless
```

These headers are what enables `crossOriginIsolated === true` in the dialog
window, which is the prerequisite for `SharedArrayBuffer` and the Pyodide SAB
bridge. Without them, no Spark queries can run.

The `deploy/` stack (static host via `halverneus/static-file-server`) sets
these headers. For custom hosting (GitHub Pages, Azure Blob, Netlify, etc.),
configure the headers at the CDN / edge layer.

### GitHub Pages

GitHub Pages can't set response headers, so use the bundled COI **service worker**
to inject them client-side. `public/vendor/coi-serviceworker.js` re-emits the
page's own responses with COOP/COEP and flips `crossOriginIsolated` to true after
a one-time automatic reload. Register it near the top of the hosted page (the
demo page already does):

```html
<script src="/vendor/coi-serviceworker.js"></script>
```

Caveats: serve at a **root** origin (a custom domain or `https://<user>.github.io`)
so the app's absolute asset paths (`/vendor/`, `/pyodide/`) resolve - a project
subpath (`/<repo>/`) needs Vite `base` configured. The service worker only covers
same-origin responses, so vendor Pyodide and the wheels same-origin (see
[reuse.md](reuse.md)).

### Azure Static Web Apps

Azure Static Web Apps supports custom headers via `staticwebapp.config.json`:

```json
{
  "globalHeaders": {
    "Cross-Origin-Opener-Policy": "same-origin",
    "Cross-Origin-Embedder-Policy": "credentialless"
  }
}
```

### Netlify

Via `netlify.toml`:

```toml
[[headers]]
  for = "/*"
  [headers.values]
    Cross-Origin-Opener-Policy = "same-origin"
    Cross-Origin-Embedder-Policy = "credentialless"
```

---

## Bring-your-own-Spark model

This add-in has **no backend server of its own**. Each user provides their own
Spark Connect endpoint. The add-in stores only:

- The endpoint host, port, and TLS flag (in `Office.context.document.settings`)
- The bearer token (in `OfficeRuntime.storage` - never in the workbook file)

What this means for distribution:
- You distribute the add-in bundle (static HTML/JS/CSS).
- Users configure their own Spark Connect endpoint in the task pane.
- There is no per-user backend registration, no cloud database, no telemetry.

---

## Production manifest

Before hosting for others:

- [ ] `manifest.xml` `SourceLocation` points to the production origin (not
      `https://localhost:3000`) - use `npm run build:manifest` below
- [ ] `AppDomains` lists the production origin
- [ ] Production bundle hosted on HTTPS with COOP/COEP headers
- [ ] `office-addin-manifest validate manifest.xml` passes
- [ ] Tested on a Chromium-based Excel host (Windows / Microsoft 365 or the web)
- [ ] The manifest's `Id` GUID is unchanged (changing it breaks existing installs)

### Updating the production origin in manifest.xml

The manifest ships targeting `https://localhost:3000` (dev). Generate a
production manifest with the bundled script - it substitutes the origin and
prints the validate command:

```bash
npm run build:manifest -- --origin https://addin.example.com
# writes dist/manifest.xml (override with --out <path>)
npx office-addin-manifest validate dist/manifest.xml
```

The origin must be a bare HTTPS origin (no trailing slash, no path); the script
rejects anything else. It replaces every `https://localhost:3000` occurrence
(SourceLocation, AppDomains, icon + resource URLs) in one pass.

---

## Sideload for enterprise (Centralized Deployment)

For Microsoft 365 tenants, IT admins can deploy the add-in to users via the
Microsoft 365 Admin Center (Centralized Deployment). Upload `manifest.xml`
and set the production SourceLocation to a URL accessible from the tenant.

Users in the target groups will see the add-in automatically in their Excel
ribbon without needing to sideload it themselves.
