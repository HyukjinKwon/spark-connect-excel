// SPDX-License-Identifier: Apache-2.0
import { defineConfig } from "vite";
import { resolve } from "node:path";

// Multi-page Office add-in:
//   - taskpane.html : the Excel task pane (UI surface, NOT cross-origin isolated)
//   - dialog.html   : the COI host window (Pyodide + pcw live here; see DECISIONS #1/#2)
//
// Cross-origin-isolation headers so the dialog page gets
// `crossOriginIsolated === true` for SharedArrayBuffer. COEP `credentialless`
// so cross-origin Pyodide/PyPI loads don't need CORP.
//
// CRITICAL: these go on the dialog (and its resources + the demo) but NOT on the
// task pane. The task pane is framed by the Office host; if it is served with
// `COOP: same-origin` it lands in its own browsing-context group, and then the
// Office Dialog API on Excel on the web refuses to open the (also isolated)
// dialog with "the dialog's domain and the add-in host's domain are not in the
// same security zone". Keeping the task pane non-isolated avoids that, while the
// dialog window it opens is still isolated and gets SharedArrayBuffer.
const coiHeaders = {
  "Cross-Origin-Opener-Policy": "same-origin",
  "Cross-Origin-Embedder-Policy": "credentialless",
};

// The task pane must NOT be cross-origin isolated; everything else (dialog,
// demo, and the same-origin assets they load) should be.
const needsCoi = (url: string | undefined): boolean =>
  !!url && !url.startsWith("/taskpane");

function coiExceptTaskpane() {
  const mw = (
    req: { url?: string },
    res: { setHeader(k: string, v: string): void },
    next: () => void,
  ) => {
    if (needsCoi(req.url)) {
      res.setHeader("Cross-Origin-Opener-Policy", coiHeaders["Cross-Origin-Opener-Policy"]);
      res.setHeader("Cross-Origin-Embedder-Policy", coiHeaders["Cross-Origin-Embedder-Policy"]);
    }
    next();
  };
  return {
    name: "coi-except-taskpane",
    configureServer(server: { middlewares: { use(fn: typeof mw): void } }) {
      server.middlewares.use(mw);
    },
    configurePreviewServer(server: { middlewares: { use(fn: typeof mw): void } }) {
      server.middlewares.use(mw);
    },
  };
}

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
    plugins: [coiExceptTaskpane()],
    server: {
      port: 3000,
      https,
      fs: {
        // Allow importing the canonical Python runtime (python/*.py) as ?raw from src/.
        allow: [resolve(__dirname)],
      },
    },
    preview: {
      port: 3000,
      https,
    },
    build: {
      outDir: resolve(__dirname, "dist"),
      emptyOutDir: true,
      target: "es2020",
      rollupOptions: {
        input: {
          taskpane: resolve(__dirname, "src/taskpane/taskpane.html"),
          dialog: resolve(__dirname, "src/dialog/dialog.html"),
          demo: resolve(__dirname, "src/demo/demo.html"),
        },
      },
    },
  };
});
