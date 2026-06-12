// SPDX-License-Identifier: Apache-2.0
//
// demo.ts - standalone, zero-install web demo of Spark Connect in the browser.
//
// Unlike the Excel add-in, this page has no Office host: it instantiates the
// runtime (Lane C PyodideHost) + the bridge (Lane D SparkBridgeHost) directly
// and renders results into an HTML table. It exposes BOTH:
//   - SQL mode    - bridge.runSQL(sql, cap) -> rich result table
//   - Python mode - host.runPython(code) -> real PySpark, `spark` pre-bound
//
// Requires the page to be cross-origin isolated (served with COOP/COEP) so the
// SharedArrayBuffer bridge works - exactly like the dialog host.

import { PyodideHost } from "../runtime/pyodideHost";
import { SparkBridgeHost } from "../bridge/sparkBridgeHost";
import { buildRemoteUri } from "../connection/connectionStore";
import { renderResultTable } from "./resultTable";
import { createCodeEditor } from "./editor";

type Mode = "sql" | "python";

// Bind the connected session to `spark` for Python-mode snippets. The runtime
// module is loaded into sys.modules by SparkBridgeHost.ensureReady().
const PY_PREAMBLE = "import spark_excel_runtime as _scx\nspark = _scx._spark\n";

const SQL_SAMPLE = "SELECT 1 AS id, 'hello' AS msg";
const PY_SAMPLE = 'spark.range(10).filter("id % 2 = 0").toPandas()';

const host = new PyodideHost();
const bridge = new SparkBridgeHost(host);
let connected = false;
let mode: Mode = "sql";

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  props: Partial<HTMLElementTagNameMap[K]> & { className?: string } = {},
  children: (Node | string)[] = [],
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  Object.assign(node, props);
  for (const c of children) node.append(typeof c === "string" ? document.createTextNode(c) : c);
  return node;
}

function isolated(): boolean {
  return typeof SharedArrayBuffer !== "undefined" && self.crossOriginIsolated === true;
}

function boot(): void {
  const app = document.getElementById("app");
  if (!app) return;
  app.replaceChildren();

  // -- Header
  app.append(
    el("div", { className: "demo-header" }, [
      el("div", { className: "demo-header__logo", textContent: "S" }),
      el("span", { className: "demo-header__title", textContent: "Spark Connect - Web Demo" }),
    ]),
    el("p", {
      className: "demo-sub",
      textContent:
        "Run Spark SQL or PySpark against your own Spark Connect cluster - entirely " +
        "in your browser, no backend. Powered by pyspark-connect-web.",
    }),
  );

  if (!isolated()) {
    app.append(
      el("div", { className: "demo-banner" }, [
        el("strong", { textContent: "This browser context isn't cross-origin isolated. " }),
        "The engine needs SharedArrayBuffer (COOP: same-origin + COEP: credentialless). " +
          "Serve this page with those headers (the dev server and the deploy stack do), " +
          "and use a Chromium-based browser.",
      ]),
    );
    return;
  }

  // -- Connection card
  const cHost = el("input", {
    className: "demo-input-host",
    value: "localhost",
    type: "text",
  }) as HTMLInputElement;
  const cPort = el("input", {
    className: "demo-input-port",
    value: "8081",
    type: "number",
  }) as HTMLInputElement;
  const cTls = el("input", { type: "checkbox" }) as HTMLInputElement;
  const cToken = el("input", {
    className: "demo-input-token",
    type: "password",
    placeholder: "(optional)",
  }) as HTMLInputElement;
  const connectBtn = el("button", {
    className: "demo-btn-secondary",
    textContent: "Connect",
  }) as HTMLButtonElement;
  const connStatus = el("div", { className: "demo-status" });

  const field = (label: string, input: HTMLElement, extraClass?: string) => {
    const d = el("div", { className: "demo-field" + (extraClass ? " " + extraClass : "") });
    d.append(el("label", { textContent: label }), input);
    return d;
  };

  app.append(
    el("div", { className: "demo-card" }, [
      el("div", { className: "demo-row" }, [
        field("Host", cHost),
        field("Port", cPort),
        field("TLS", cTls, "demo-field--tls"),
        field("Token", cToken),
        connectBtn,
      ]),
      connStatus,
    ]),
  );

  // -- Query card: mode toggle, max rows, run button, code editor, status

  // Mode toggle (segmented control - neutral styling, not orange).
  const sqlBtn = el("button", { textContent: "SQL" }) as HTMLButtonElement;
  const pyBtn = el("button", { textContent: "Python" }) as HTMLButtonElement;
  const modeToggle = el("div", { className: "demo-mode-toggle" }, [sqlBtn, pyBtn]);

  // Max-rows input.
  const capInput = el("input", {
    className: "demo-input-cap",
    type: "number",
    value: "1000",
  }) as HTMLInputElement;

  // Run button - blue, clearly the primary action.
  const runBtn = el("button", {
    className: "demo-btn-run",
    textContent: "Run",
  }) as HTMLButtonElement;

  const runStatus = el("div", { className: "demo-status" });
  const results = el("div", {});

  // Code editor (syntax-highlighted overlay).
  const codeEditor = createCodeEditor({ value: SQL_SAMPLE, mode: "sql" });

  function applyMode(next: Mode): void {
    mode = next;
    // Update segmented control appearance.
    sqlBtn.className = next === "sql" ? "is-active" : "";
    pyBtn.className = next === "python" ? "is-active" : "";
    // Swap snippet and re-highlight.
    codeEditor.setValue(next === "sql" ? SQL_SAMPLE : PY_SAMPLE);
    codeEditor.setMode(next);
  }

  sqlBtn.onclick = () => applyMode("sql");
  pyBtn.onclick = () => applyMode("python");

  app.append(
    el("div", { className: "demo-card" }, [
      el("div", { className: "demo-row" }, [
        el("div", { className: "demo-field" }, [el("label", { textContent: "Mode" }), modeToggle]),
        field("Max rows (SQL)", capInput),
        runBtn,
      ]),
      codeEditor.el,
      runStatus,
    ]),
    results,
  );

  // Initial mode render.
  applyMode("sql");

  // -- Wiring
  function busy(b: boolean, statusEl: HTMLElement, msg = ""): void {
    runBtn.disabled = b;
    connectBtn.disabled = b;
    statusEl.classList.remove("is-error");
    statusEl.textContent = msg;
  }
  function fail(statusEl: HTMLElement, err: unknown): void {
    statusEl.classList.add("is-error");
    statusEl.textContent = err instanceof Error ? err.message : String(err);
  }

  connectBtn.onclick = async () => {
    busy(true, connStatus, "Starting engine and connecting... (first run downloads Pyodide)");
    try {
      const uri = buildRemoteUri({
        host: cHost.value.trim(),
        port: Number(cPort.value) || 8081,
        tls: cTls.checked,
      });
      await bridge.connect(uri, cToken.value ? { token: cToken.value } : undefined);
      connected = true;
      connStatus.textContent = `Connected - ${uri}`;
    } catch (err) {
      fail(connStatus, err);
    } finally {
      runBtn.disabled = false;
      connectBtn.disabled = false;
    }
  };

  runBtn.onclick = async () => {
    const src = codeEditor.getValue().trim();
    if (!src) return;
    busy(true, runStatus, "Running...");
    try {
      if (mode === "sql") {
        const cap = Number(capInput.value) || 1000;
        const result = await bridge.runSQL(src, cap);
        results.replaceChildren(renderResultTable(result));
        runStatus.textContent = "Done.";
      } else {
        if (!connected) await bridge.ensureReady();
        const out = await host.runPython(PY_PREAMBLE + src);
        results.replaceChildren(el("pre", { className: "demo-table", textContent: out }));
        runStatus.textContent = "Done. (print() output goes to the browser console.)";
      }
    } catch (err) {
      fail(runStatus, err);
    } finally {
      runBtn.disabled = false;
      connectBtn.disabled = false;
    }
  };

  app.append(
    el("div", { className: "demo-footer" }, [
      "Powered by ",
      el("a", {
        href: "https://github.com/HyukjinKwon/pyspark-client-wasm",
        textContent: "pyspark-connect-web",
      }),
      ". You connect to your own Spark Connect cluster; nothing is sent to a third party.",
    ]),
  );
}

boot();
