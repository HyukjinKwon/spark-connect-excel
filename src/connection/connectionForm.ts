// SPDX-License-Identifier: Apache-2.0
//
// connectionForm.ts — pure render/validate helpers for the connection UI.
//
// This module has NO global side effects and NO direct Office API calls.
// It is designed to be composed by Lane E (Query UX) which owns the task-pane
// entry point. All DOM work uses document.createElement (no framework).
//
// Exports:
//   renderConnectionForm(root, initial, onSubmit)  — renders the connection form
//   validateConnection(host, port)                 — returns error string or null
//   connectionStatusBadge(root, status)            — renders the RuntimeStatus flags

import type { ConnectionConfig } from "./connectionStore.js";
import { buildRemoteUri } from "./connectionStore.js";
import { SC_URI_HINT } from "../seam.js";
import type { RuntimeStatus } from "../seam.js";

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

/**
 * Validate raw user input for the host and port fields.
 *
 * Returns a human-readable error message string when validation fails,
 * or `null` when the inputs are acceptable.
 *
 * A valid host is a non-empty string that is not pure whitespace.
 * A valid port is an integer in the range [1, 65535].
 */
export function validateConnection(host: string, port: string | number): string | null {
  const trimmedHost = typeof host === "string" ? host.trim() : "";
  if (trimmedHost.length === 0) {
    return "Host is required.";
  }

  const portNum = typeof port === "number" ? port : parseInt(port, 10);
  if (!Number.isInteger(portNum) || portNum < 1 || portNum > 65535) {
    return "Port must be a whole number between 1 and 65535.";
  }

  return null;
}

// ---------------------------------------------------------------------------
// Connection form
// ---------------------------------------------------------------------------

export interface ConnectionFormOptions {
  /** If true, show the bearer token field. Default: true. */
  showToken?: boolean;
}

/**
 * Render a connection form into `root`, replacing any existing children.
 *
 * The form contains:
 *   - Host input (text)
 *   - Port input (number, default 8081)
 *   - TLS checkbox
 *   - Bearer token input (password, optional — hidden by default for clarity)
 *   - A read-only URI preview that updates as the user types
 *   - A "Connect" submit button
 *
 * `initial` pre-fills the fields when provided (e.g. loaded from storage).
 * `onSubmit` is called with the validated `ConnectionConfig` and optional token
 * when the user submits; invalid input shows inline error text instead.
 *
 * No global state is mutated. Returns a `dispose` function that removes all
 * event listeners (call when the form is torn down).
 */
export function renderConnectionForm(
  root: HTMLElement,
  initial: ConnectionConfig | null,
  onSubmit: (cfg: ConnectionConfig, token?: string) => void,
  opts: ConnectionFormOptions = {},
): () => void {
  const showToken = opts.showToken !== false; // default true

  // Clear existing children.
  while (root.firstChild) root.removeChild(root.firstChild);

  // ---- host ---------------------------------------------------------------
  const hostLabel = el("label", { htmlFor: "sc-host" }, "Spark Connect host");
  const hostInput = el("input", {
    id: "sc-host",
    type: "text",
    placeholder: "localhost",
    value: initial?.host ?? "",
    autocomplete: "off",
    spellcheck: "false",
  }) as HTMLInputElement;

  // ---- port ---------------------------------------------------------------
  const portLabel = el("label", { htmlFor: "sc-port" }, "Port");
  const portInput = el("input", {
    id: "sc-port",
    type: "number",
    min: "1",
    max: "65535",
    value: String(initial?.port ?? 8081),
  }) as HTMLInputElement;

  // ---- TLS ----------------------------------------------------------------
  const tlsLabel = el("label", { htmlFor: "sc-tls" }, "Use TLS");
  const tlsInput = el("input", {
    id: "sc-tls",
    type: "checkbox",
  }) as HTMLInputElement;
  tlsInput.checked = initial?.tls ?? false;

  // ---- URI preview --------------------------------------------------------
  const uriPreviewLabel = el("label", {}, "Connection URI (auto)");
  const uriPreview = el("input", {
    type: "text",
    readOnly: "true",
    disabled: "true",
    placeholder: SC_URI_HINT,
    title: "This URI is built automatically from host, port, and TLS settings.",
    style: "font-family: monospace; font-size: 0.85em; color: #555;",
  }) as HTMLInputElement;

  function refreshUriPreview(): void {
    const host = hostInput.value.trim();
    const portRaw = portInput.value.trim();
    const portNum = parseInt(portRaw, 10);
    if (host && Number.isInteger(portNum) && portNum >= 1 && portNum <= 65535) {
      const cfg: ConnectionConfig = { host, port: portNum, tls: tlsInput.checked };
      uriPreview.value = buildRemoteUri(cfg);
    } else {
      uriPreview.value = "";
      uriPreview.placeholder = SC_URI_HINT;
    }
  }

  hostInput.addEventListener("input", refreshUriPreview);
  portInput.addEventListener("input", refreshUriPreview);
  tlsInput.addEventListener("change", refreshUriPreview);
  refreshUriPreview();

  // ---- bearer token (optional) -------------------------------------------
  // DECISIONS #6: the token is shown/collected in the UI but NEVER stored in
  // document settings or a cell. connectionStore.saveToken() handles the safe
  // storage path; this form just captures the raw string and hands it to the
  // onSubmit callback.
  let tokenInput: HTMLInputElement | null = null;
  let tokenLabel: HTMLElement | null = null;
  if (showToken) {
    tokenLabel = el("label", { htmlFor: "sc-token" }, "Bearer token (optional)");
    tokenInput = el("input", {
      id: "sc-token",
      type: "password",
      placeholder: "Leave blank if the cluster does not require auth",
      autocomplete: "off",
    }) as HTMLInputElement;
  }

  // ---- error display ------------------------------------------------------
  const errorEl = el(
    "p",
    {
      role: "alert",
      style: "color: #c00; min-height: 1.2em; margin: 0;",
    },
    "",
  );

  // ---- submit button -------------------------------------------------------
  const submitBtn = el("button", { type: "submit" }, "Connect") as HTMLButtonElement;

  // ---- form submit handler ------------------------------------------------
  function handleSubmit(e: Event): void {
    e.preventDefault();
    const host = hostInput.value.trim();
    const portRaw = portInput.value.trim();
    const portNum = parseInt(portRaw, 10);

    const err = validateConnection(host, portRaw);
    if (err) {
      errorEl.textContent = err;
      hostInput.focus();
      return;
    }
    errorEl.textContent = "";

    const cfg: ConnectionConfig = { host, port: portNum, tls: tlsInput.checked };
    const token = tokenInput?.value.trim() || undefined;
    onSubmit(cfg, token);
  }

  // ---- assemble form ------------------------------------------------------
  const form = el("form", { autocomplete: "off" }) as HTMLFormElement;
  form.addEventListener("submit", handleSubmit);

  function row(...children: HTMLElement[]): HTMLElement {
    const div = el("div", { style: "display:flex; flex-direction:column; margin-bottom:8px;" });
    children.forEach((c) => div.appendChild(c));
    return div;
  }

  form.appendChild(row(hostLabel, hostInput));
  form.appendChild(row(portLabel, portInput));

  // TLS in a horizontal row (checkbox + label side by side)
  const tlsRow = el("div", {
    style: "display:flex; align-items:center; gap:6px; margin-bottom:8px;",
  });
  tlsRow.appendChild(tlsInput);
  tlsRow.appendChild(tlsLabel);
  form.appendChild(tlsRow);

  form.appendChild(row(uriPreviewLabel, uriPreview));

  if (showToken && tokenLabel && tokenInput) {
    form.appendChild(row(tokenLabel, tokenInput));
  }

  form.appendChild(errorEl);
  form.appendChild(submitBtn);
  root.appendChild(form);

  // Return a dispose function that cleans up listeners.
  return function dispose(): void {
    hostInput.removeEventListener("input", refreshUriPreview);
    portInput.removeEventListener("input", refreshUriPreview);
    tlsInput.removeEventListener("change", refreshUriPreview);
    form.removeEventListener("submit", handleSubmit);
  };
}

// ---------------------------------------------------------------------------
// RuntimeStatus badge
// ---------------------------------------------------------------------------

/**
 * Render (or update) a compact status badge inside `root` that displays the
 * three `RuntimeStatus` flags: `crossOriginIsolated`, `pyodideReady`, and
 * `connected`.
 *
 * Calling this function multiple times on the same `root` is idempotent — it
 * replaces the existing badge rather than appending a new one.
 */
export function connectionStatusBadge(root: HTMLElement, status: RuntimeStatus): void {
  // Find or create the badge container.
  let badge = root.querySelector<HTMLElement>("[data-sc-status-badge]");
  if (!badge) {
    badge = el("div", {
      "data-sc-status-badge": "",
      style: "display:flex; gap:10px; font-size:0.8em; padding:4px 0;",
      role: "status",
      "aria-live": "polite",
    });
    root.appendChild(badge);
  }

  // Rebuild flags.
  while (badge.firstChild) badge.removeChild(badge.firstChild);

  const flags: Array<[keyof RuntimeStatus, string]> = [
    ["crossOriginIsolated", "COI"],
    ["pyodideReady", "Pyodide"],
    ["connected", "Connected"],
  ];

  for (const [key, label] of flags) {
    const ok = status[key];
    const flag = el(
      "span",
      {
        title: key,
        style: [
          "display:inline-flex; align-items:center; gap:3px;",
          "padding:2px 6px; border-radius:4px;",
          ok ? "background:#d4f5d4; color:#1a6b1a;" : "background:#f5d4d4; color:#6b1a1a;",
        ].join(" "),
      },
      `${ok ? "✓" : "✗"} ${label}`,
    );
    badge.appendChild(flag);
  }
}

// ---------------------------------------------------------------------------
// Internal: tiny createElement helper — no framework
// ---------------------------------------------------------------------------

function el(tag: string, attrs: Record<string, string | boolean> = {}, text?: string): HTMLElement {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "htmlFor") {
      (node as HTMLLabelElement).htmlFor = String(v);
    } else if (typeof v === "boolean") {
      if (v) node.setAttribute(k, "");
      // false → omit
    } else {
      node.setAttribute(k, v);
    }
  }
  if (text !== undefined) node.textContent = text;
  return node;
}
