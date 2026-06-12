// SPDX-License-Identifier: Apache-2.0
//
// demo.ts — standalone, zero-install web demo of Spark Connect in the browser.
//
// Unlike the Excel add-in, this page has no Office host: it instantiates the
// runtime (Lane C PyodideHost) + the bridge (Lane D SparkBridgeHost) directly
// and renders results into an HTML table. It exposes BOTH:
//   - SQL mode    — bridge.runSQL(sql, cap) -> rich result table
//   - Python mode — host.runPython(code) -> real PySpark, `spark` pre-bound
//
// Requires the page to be cross-origin isolated (served with COOP/COEP) so the
// SharedArrayBuffer bridge works — exactly like the dialog host.

import { PyodideHost } from "../runtime/pyodideHost";
import { SparkBridgeHost } from "../bridge/sparkBridgeHost";
import { buildRemoteUri } from "../connection/connectionStore";
import { renderResultTable } from "./resultTable";

type Mode = "sql" | "python";

// Bind the connected session to `spark` for Python-mode snippets. The runtime
// module is loaded into sys.modules by SparkBridgeHost.ensureReady().
const PY_PREAMBLE = "import spark_excel_runtime as _scx\nspark = _scx._spark\n";

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

  // ── Header
  app.append(
    el("div", { className: "demo-header" }, [
      el("div", { className: "demo-header__logo", textContent: "⚡" }),
      el("span", { className: "demo-header__title", textContent: "Spark Connect — Web Demo" }),
    ]),
    el("p", {
      className: "demo-sub",
      textContent:
        "Run Spark SQL or PySpark against your own Spark Connect cluster — entirely " +
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

  // ── Connection card
  const cHost = el("input", { value: "localhost", size: 16 }) as HTMLInputElement;
  const cPort = el("input", { value: "8081", size: 6, type: "number" }) as HTMLInputElement;
  const cTls = el("input", { type: "checkbox" }) as HTMLInputElement;
  const cToken = el("input", {
    type: "password",
    size: 18,
    placeholder: "(optional)",
  }) as HTMLInputElement;
  const connectBtn = el("button", {
    className: "demo-btn-secondary",
    textContent: "Connect",
  }) as HTMLButtonElement;
  const connStatus = el("div", { className: "demo-status" });

  const field = (label: string, input: HTMLElement) =>
    el("div", { className: "demo-field" }, [el("label", { textContent: label }), input]);

  app.append(
    el("div", { className: "demo-card" }, [
      el("div", { className: "demo-row" }, [
        field("Host", cHost),
        field("Port", cPort),
        field("TLS", cTls),
        field("Token", cToken),
        connectBtn,
      ]),
      connStatus,
    ]),
  );

  // ── Query card (mode toggle + editor + run)
  const sqlBtn = el("button", {
    className: "demo-btn-secondary",
    textContent: "SQL",
  }) as HTMLButtonElement;
  const pyBtn = el("button", {
    className: "demo-btn-secondary",
    textContent: "Python",
  }) as HTMLButtonElement;
  const editor = el("textarea", {
    className: "demo-textarea",
    value: "SELECT 1 AS id, 'hello' AS msg",
  }) as HTMLTextAreaElement;
  const capWrap = el("input", { type: "number", value: "1000", size: 6 }) as HTMLInputElement;
  const runBtn = el("button", {
    className: "demo-btn-primary",
    textContent: "▶ Run",
  }) as HTMLButtonElement;
  const runStatus = el("div", { className: "demo-status" });
  const results = el("div", {});

  function setMode(next: Mode): void {
    mode = next;
    sqlBtn.className = next === "sql" ? "demo-btn-primary" : "demo-btn-secondary";
    pyBtn.className = next === "python" ? "demo-btn-primary" : "demo-btn-secondary";
    editor.value =
      next === "sql"
        ? "SELECT 1 AS id, 'hello' AS msg"
        : 'spark.range(10).filter("id % 2 = 0").toPandas()';
  }
  sqlBtn.onclick = () => setMode("sql");
  pyBtn.onclick = () => setMode("python");

  app.append(
    el("div", { className: "demo-card" }, [
      el("div", { className: "demo-row" }, [
        el("div", { className: "demo-field" }, [
          el("label", { textContent: "Mode" }),
          el("div", { className: "demo-row" }, [sqlBtn, pyBtn]),
        ]),
        field("Max rows (SQL)", capWrap),
        runBtn,
      ]),
      editor,
      runStatus,
    ]),
    results,
  );
  setMode("sql");

  // ── Wiring
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
    busy(true, connStatus, "Starting engine and connecting… (first run downloads Pyodide)");
    try {
      const uri = buildRemoteUri({
        host: cHost.value.trim(),
        port: Number(cPort.value) || 8081,
        tls: cTls.checked,
      });
      await bridge.connect(uri, cToken.value ? { token: cToken.value } : undefined);
      connected = true;
      connStatus.textContent = `Connected — ${uri}`;
    } catch (err) {
      fail(connStatus, err);
    } finally {
      runBtn.disabled = false;
      connectBtn.disabled = false;
    }
  };

  runBtn.onclick = async () => {
    const src = editor.value.trim();
    if (!src) return;
    busy(true, runStatus, "Running…");
    try {
      if (mode === "sql") {
        const cap = Number(capWrap.value) || 1000;
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
