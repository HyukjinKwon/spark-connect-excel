// SPDX-License-Identifier: Apache-2.0
//
// taskpane.ts — Task-pane entry point (Lane E).
//
// Bootstraps Office.js, creates the COI dialog bridge (Lane D), and delegates
// all UI to queryPanel.ts.  This file stays thin: lifecycle + wiring only.
//
// DECISIONS #7: MUST NOT import from src/runtime/**.

import { createDialogBridge } from "../bridge/sparkBridgeClient.js";
import type { BridgeEvent } from "../seam.js";
import { renderQueryPanel } from "./queryPanel.js";

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------

Office.onReady(() => {
  boot().catch((err: unknown) => {
    const msg = err instanceof Error ? err.message : String(err);
    renderFatalError(`Failed to start: ${msg}`);
  });
});

async function boot(): Promise<void> {
  const appRoot = document.getElementById("app");
  if (!appRoot) throw new Error("#app not found");

  // Build the pane chrome (header + scrollable body).
  const pane = buildPaneChrome(appRoot);

  // Show an engine-starting banner while the dialog is being opened.
  const engineBanner = buildEngineBanner("Opening Spark engine window…");
  pane.body.appendChild(engineBanner);

  // ── Open the COI dialog window (Lane D / Lane B).
  // The dialog URL is resolved relative to the add-in origin so it works on
  // all Excel hosts (web, Windows WebView2, Mac WKWebView).
  const dialogUrl = new URL("/dialog/dialog.html", location.origin).href;

  let bridge;
  try {
    bridge = await createDialogBridge(dialogUrl, {
      width: 60,
      height: 60,
      onEvent: (evt: BridgeEvent) => {
        // Unsupported host (no cross-origin isolation / SharedArrayBuffer):
        // replace the UI with a blocking guidance panel.
        if (evt.event === "unsupported") {
          removeBanner(engineBanner);
          const reason =
            (evt.payload as { reason?: string } | undefined)?.reason ??
            "This Excel host can't run the Spark engine.";
          renderUnsupportedInBody(pane.body, reason);
          return;
        }
        // Forward progress events to the banner; forward status events to the
        // query panel's refreshStatus hook (set after renderQueryPanel runs).
        if (evt.event === "progress" && typeof evt.payload === "string") {
          updateEngineBanner(engineBanner, evt.payload);
        }
        // Status events — refreshStatus may not exist yet during early boot.
        if (evt.event === "status" || evt.event === "ready") {
          const panelApi = getPanelApi(pane.body);
          if (panelApi) panelApi.refreshStatus();
        }
      },
    });
  } catch (err) {
    removeBanner(engineBanner);
    renderErrorInBody(pane.body, `Could not open engine window: ${errMsg(err)}`);
    return;
  }

  // ── Start Pyodide + pcw in the background; update the banner as it boots.
  updateEngineBanner(engineBanner, "Starting Spark engine…");

  bridge
    .ensureReady()
    .then(() => {
      removeBanner(engineBanner);
      const panelApi = getPanelApi(pane.body);
      if (panelApi) panelApi.refreshStatus();
    })
    .catch((err: unknown) => {
      updateEngineBanner(engineBanner, `Engine error: ${errMsg(err)}`, true);
    });

  // ── Render query panel — doesn't wait for ensureReady; user can configure
  // the connection while Pyodide loads in the background.
  removeBanner(engineBanner);

  const { api } = renderQueryPanel(pane.body, bridge);

  // Stash the api so the onEvent callback above can call refreshStatus.
  setPanelApi(pane.body, api);
}

// ---------------------------------------------------------------------------
// Pane chrome builder
// ---------------------------------------------------------------------------

interface PaneChrome {
  pane: HTMLDivElement;
  header: HTMLDivElement;
  body: HTMLDivElement;
}

function buildPaneChrome(root: HTMLElement): PaneChrome {
  while (root.firstChild) root.removeChild(root.firstChild);

  const pane = document.createElement("div");
  pane.className = "sc-pane";

  // Header bar with logo + title.
  const header = document.createElement("div");
  header.className = "sc-header";

  const logo = document.createElement("div");
  logo.className = "sc-header__logo";
  logo.textContent = "S";
  logo.setAttribute("aria-hidden", "true");

  const title = document.createElement("span");
  title.className = "sc-header__title";
  title.textContent = "Spark Connect for Excel";

  header.appendChild(logo);
  header.appendChild(title);

  const body = document.createElement("div");
  body.className = "sc-body";

  pane.appendChild(header);
  pane.appendChild(body);
  root.appendChild(pane);

  return { pane, header, body };
}

// ---------------------------------------------------------------------------
// Engine-starting banner
// ---------------------------------------------------------------------------

function buildEngineBanner(msg: string): HTMLDivElement {
  const banner = document.createElement("div");
  banner.className = "sc-engine-banner";
  banner.setAttribute("role", "status");
  banner.setAttribute("aria-live", "polite");

  const spinner = document.createElement("div");
  spinner.className = "sc-engine-banner__spinner";
  spinner.setAttribute("aria-hidden", "true");

  const text = document.createElement("span");
  text.textContent = msg;

  banner.appendChild(spinner);
  banner.appendChild(text);
  return banner;
}

function updateEngineBanner(banner: HTMLDivElement, msg: string, isError = false): void {
  const textEl = banner.querySelector("span");
  if (textEl) textEl.textContent = msg;
  if (isError) {
    banner.style.background = "#fff5f5";
    banner.style.borderColor = "#f5c6cb";
    banner.style.color = "#721c24";
    const spinner = banner.querySelector<HTMLElement>(".sc-engine-banner__spinner");
    if (spinner) spinner.style.display = "none";
  }
}

function removeBanner(banner: HTMLDivElement): void {
  banner.parentNode?.removeChild(banner);
}

// ---------------------------------------------------------------------------
// Panel API stash (keyed on the body element — no global state)
// ---------------------------------------------------------------------------

const _panelApiMap = new WeakMap<HTMLElement, { refreshStatus(): void }>();

function setPanelApi(body: HTMLElement, api: { refreshStatus(): void }): void {
  _panelApiMap.set(body, api);
}

function getPanelApi(body: HTMLElement): { refreshStatus(): void } | undefined {
  return _panelApiMap.get(body);
}

// ---------------------------------------------------------------------------
// Error display helpers
// ---------------------------------------------------------------------------

function renderFatalError(msg: string): void {
  const app = document.getElementById("app");
  if (!app) return;
  while (app.firstChild) app.removeChild(app.firstChild);
  const errEl = document.createElement("div");
  errEl.className = "sc-error";
  errEl.style.margin = "16px";
  errEl.textContent = msg;
  app.appendChild(errEl);
}

function renderErrorInBody(body: HTMLElement, msg: string): void {
  const errEl = document.createElement("div");
  errEl.className = "sc-error";
  errEl.textContent = msg;
  body.appendChild(errEl);
}

/** Replace the body with a clear, blocking "unsupported host" guidance panel. */
function renderUnsupportedInBody(body: HTMLElement, reason: string): void {
  while (body.firstChild) body.removeChild(body.firstChild);

  const panel = document.createElement("div");
  panel.className = "sc-unsupported";

  const heading = document.createElement("h2");
  heading.className = "sc-unsupported__title";
  heading.textContent = "Unsupported Excel host";

  const detail = document.createElement("p");
  detail.className = "sc-unsupported__detail";
  detail.textContent = reason;

  const list = document.createElement("ul");
  list.className = "sc-unsupported__list";
  for (const item of [
    "Excel on Windows (Microsoft 365)",
    "Excel on the web in Microsoft Edge or Google Chrome",
  ]) {
    const li = document.createElement("li");
    li.textContent = item;
    list.appendChild(li);
  }

  panel.appendChild(heading);
  panel.appendChild(detail);
  const supported = document.createElement("p");
  supported.className = "sc-unsupported__detail";
  supported.textContent = "Supported hosts:";
  panel.appendChild(supported);
  panel.appendChild(list);
  body.appendChild(panel);
}

function errMsg(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
