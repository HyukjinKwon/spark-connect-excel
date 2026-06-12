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

export default defineConfig({
  root: "src",
  publicDir: resolve(__dirname, "public"),
  server: {
    port: 3000,
    https: undefined, // office-addin-dev-certs wires HTTPS in `npm run dev:https`; plain for CI
    headers: coiHeaders,
    fs: {
      // Allow importing the canonical Python runtime (python/*.py) as ?raw from src/.
      allow: [resolve(__dirname)],
    },
  },
  preview: {
    port: 3000,
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
});
