<!-- SPDX-License-Identifier: Apache-2.0 -->

# Distribution

## Two distribution paths

### 1. Sideload (developer / self-hosted / private team)

Sideloading installs the add-in directly from a manifest URL without going
through AppSource. It is the recommended path for:

- Development and testing
- Private team deployments
- IT-managed enterprise rollouts via Centralized Deployment

See `scripts/sideload.md` and `docs/installation.md` for step-by-step instructions.

### 2. AppSource (public distribution)

AppSource submission makes the add-in available to any Microsoft 365 user via
**Insert > Get Add-ins**. This path requires:

- A Microsoft Partner Center account
- The add-in bundle hosted on a publicly accessible HTTPS server with the
  correct COI headers (see below)
- Passing Microsoft's automated and manual validation checklist

---

## Hosting requirements

Whether sideloaded or AppSource-published, the add-in bundle (`dist/`) **must**
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

### GitHub Pages example

GitHub Pages does not allow custom response headers for individual paths. Use a
CDN proxy (Cloudflare Workers / AWS CloudFront custom headers policy) in front of
GitHub Pages to inject the COI headers.

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
- The bearer token (in `OfficeRuntime.storage` — never in the workbook file)

What this means for distribution:
- You distribute the add-in bundle (static HTML/JS/CSS).
- Users configure their own Spark Connect endpoint in the task pane.
- There is no per-user backend registration, no cloud database, no telemetry.

---

## AppSource submission checklist

Before submitting to AppSource:

- [ ] `manifest.xml` updated with the production `SourceLocation` URL (not
      `https://localhost:3000`)
- [ ] `manifest.xml` version incremented
- [ ] Production bundle hosted on HTTPS with COOP/COEP headers
- [ ] `AppDomains` in `manifest.xml` lists the production origin
- [ ] `SupportUrl` points to a real page
- [ ] Icon assets (`icon-16.png`, `icon-32.png`, `icon-80.png`) provided in
      `public/assets/`
- [ ] `office-addin-manifest validate manifest.xml` passes
- [ ] Add-in tested on Windows (WebView2), Mac (WKWebView), and Excel on the web
- [ ] Privacy policy URL added to `manifest.xml` (required for AppSource)
- [ ] The manifest's `Id` GUID matches the one in source control (do not
      regenerate — changing it breaks existing user installations)

### Updating the production origin in manifest.xml

The manifest ships targeting `https://localhost:3000` (dev). Generate a
production manifest with the bundled script — it substitutes the origin and
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
