// SPDX-License-Identifier: Apache-2.0
import { defineConfig } from "vite";
import { resolve } from "node:path";

// Multi-page Office add-in:
//   - taskpane.html : the Excel task pane (UI surface, NOT cross-origin isolated)
//   - dialog.html   : the COI host window (Pyodide + pcw live here; see DECISIONS #1/#2)
//
// Dev server sets cross-origin-isolation headers so the dialog page (served from
// the same dev origin) gets `crossOriginIsolated === true` for SharedArrayBuffer.
// We use COEP `credentialless` so cross-origin Pyodide/PyPI loads don't need CORP.
const coiHeaders = {
  "Cross-Origin-Opener-Policy": "same-origin",
  "Cross-Origin-Embedder-Policy": "credentialless",
};

// Office add-ins require HTTPS to load the task pane on Windows (WebView2) and
// Mac (WKWebView). `npm run dev:https` sets HTTPS=true; we then serve with the
// OS-trusted cert from `office-addin-dev-certs` (run `npx office-addin-dev-certs
// install` once). Plain HTTP stays the default for unit tests, the COI e2e gate,
// and CI — none of which load inside a real Office host.
async function httpsOptions(): Promise<{ key: Buffer; cert: Buffer } | undefined> {
  if (process.env.HTTPS !== "true") return undefined;
  const mod: any = await import("office-addin-dev-certs");
  const getOpts = mod.getHttpsServerOptions ?? mod.default?.getHttpsServerOptions;
  const opts = await getOpts();
  return { key: opts.key, cert: opts.cert };
}

export default defineConfig(async () => {
  const https = await httpsOptions();
  return {
    root: "src",
    publicDir: resolve(__dirname, "public"),
    server: {
      port: 3000,
      https,
      headers: coiHeaders,
      fs: {
        // Allow importing the canonical Python runtime (python/*.py) as ?raw from src/.
        allow: [resolve(__dirname)],
      },
    },
    preview: {
      port: 3000,
      https,
      headers: coiHeaders,
    },
    build: {
      outDir: resolve(__dirname, "dist"),
      emptyOutDir: true,
      target: "es2020",
      rollupOptions: {
        input: {
          taskpane: resolve(__dirname, "src/taskpane/taskpane.html"),
          dialog: resolve(__dirname, "src/dialog/dialog.html"),
        },
      },
    },
  };
});
