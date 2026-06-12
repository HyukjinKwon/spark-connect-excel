// SPDX-License-Identifier: Apache-2.0
//
// dialogHost.ts — Office Dialog window entry point (Lane B).
//
// This module is loaded inside the Office Dialog window (a real top-level
// window we control), NOT in the task pane.  It:
//
//   1. Verifies cross-origin isolation (SharedArrayBuffer/Atomics gate).
//   2. Constructs the Pyodide runtime host (Lane C: PyodideHost) and the
//      SparkBridge implementation (Lane D: SparkBridgeHost).
//   3. Wires the two-way Office dialog message channel:
//        parent → dialog : Office.EventType.DialogParentMessageReceived
//                          → decoded BridgeRequest → bridge method dispatch
//                          → BridgeResponse via Office.context.ui.messageParent()
//        dialog → parent : BridgeEvent pushes via messageParent()
//
// Office API requirement: DialogApi 1.2 (needed for DialogParentMessageReceived
// on the dialog side and messageParent on the dialog side; see findings doc).

import type { BridgeMethod, BridgeRequest, BridgeResponse, BridgeEvent } from "../seam";
import { decodeMessage, encodeMessage } from "../seam";

// Lane C — concrete runtime host (dialog-side Pyodide boot + runPython).
// If PyodideHost is not yet written, the integrator resolves the symbol gap.
import { PyodideHost } from "../runtime/pyodideHost";

// Lane D — concrete SparkBridge implementation running inside the dialog.
// If SparkBridgeHost is not yet written, the integrator resolves the symbol gap.
import { SparkBridgeHost } from "../bridge/sparkBridgeHost";

import { ensureCrossOriginIsolated, isolationDiagnostics } from "./coi";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Update the visible #status div (best-effort — may not exist during tests). */
function setStatus(text: string, isError = false): void {
  const el = document.getElementById("status");
  if (!el) return;
  el.textContent = text;
  if (isError) {
    el.classList.add("error");
  } else {
    el.classList.remove("error");
  }
}

/** Push a BridgeEvent to the parent task pane. */
function pushEvent(evt: BridgeEvent): void {
  try {
    Office.context.ui.messageParent(encodeMessage(evt));
  } catch (err) {
    console.error("[dialogHost] messageParent(event) failed:", err);
  }
}

/** Send a BridgeResponse to the parent task pane. */
function sendResponse(res: BridgeResponse): void {
  try {
    Office.context.ui.messageParent(encodeMessage(res));
  } catch (err) {
    console.error("[dialogHost] messageParent(response) failed:", err);
  }
}

// ---------------------------------------------------------------------------
// Main boot sequence (runs inside Office.onReady)
// ---------------------------------------------------------------------------

async function main(): Promise<void> {
  // --- Step 1: COI gate ---
  setStatus("Checking cross-origin isolation…");

  const isolated = await ensureCrossOriginIsolated();
  const diag = isolationDiagnostics();

  if (!isolated) {
    const msg =
      "Cross-origin isolation is not available. " +
      "SharedArrayBuffer is required for the Spark engine but is blocked by " +
      "the browser. Ensure the dialog is served with COOP: same-origin and " +
      "COEP: credentialless (or credentialless mode is unsupported by this " +
      "browser). diagnostics=" +
      JSON.stringify(diag);

    setStatus(msg, /* isError */ true);

    // Inform the parent so it can surface a meaningful error to the user.
    pushEvent({
      kind: "evt",
      event: "log",
      payload: { level: "error", message: msg, diagnostics: diag },
    });

    // No point continuing — the runtime requires SAB.
    return;
  }

  pushEvent({
    kind: "evt",
    event: "progress",
    payload: { message: "Cross-origin isolation confirmed.", diagnostics: diag },
  });

  // --- Step 2: Construct the runtime host and bridge ---
  setStatus("Initialising Spark engine…");

  const host = new PyodideHost();
  const bridge = new SparkBridgeHost(host);

  // --- Step 3: Wire the Office dialog message channel ---
  //
  // Parent → dialog: Office.EventType.DialogParentMessageReceived
  //   (requires DialogApi 1.2; available in Excel 2016+, Office on the web,
  //    and newer shared runtime builds)
  //
  // Each incoming message is a JSON-encoded BridgeRequest; we dispatch to the
  // appropriate bridge method and reply with a BridgeResponse (same `id`).

  // Method dispatch table — keeps routing generic over BridgeMethod.
  type DispatchFn = (...args: unknown[]) => Promise<unknown> | unknown;

  const methodTable: Record<BridgeMethod, DispatchFn> = {
    ensureReady: () => bridge.ensureReady(),
    connect: (uri: unknown, opts?: unknown) =>
      bridge.connect(uri as string, opts as { token?: string } | undefined),
    runSQL: (sql: unknown, rowCap: unknown) => bridge.runSQL(sql as string, rowCap as number),
    schemaOf: (sql: unknown) => bridge.schemaOf(sql as string),
    // cancel() is synchronous on SparkBridge; we still wrap it uniformly.
    cancel: () => {
      bridge.cancel();
    },
  };

  const handleParentMessage = async (
    arg: Office.DialogParentMessageReceivedEventArgs,
  ): Promise<void> => {
    let req: BridgeRequest;
    try {
      const parsed = decodeMessage(arg.message);
      if (parsed.kind !== "req") {
        // Unexpected message shape — log and ignore.
        console.warn("[dialogHost] Received non-request message:", parsed);
        return;
      }
      req = parsed as BridgeRequest;
    } catch (parseErr) {
      console.error("[dialogHost] Failed to parse parent message:", parseErr, arg.message);
      return;
    }

    const { id, method, args } = req;

    const dispatchFn = methodTable[method];
    if (!dispatchFn) {
      sendResponse({
        kind: "res",
        id,
        ok: false,
        error: { name: "UnknownMethod", message: `Unknown bridge method: ${method}` },
      });
      return;
    }

    try {
      const result = await Promise.resolve(dispatchFn(...args));
      sendResponse({ kind: "res", id, ok: true, result: result ?? null });
    } catch (callErr) {
      const err = callErr instanceof Error ? callErr : new Error(String(callErr));
      sendResponse({
        kind: "res",
        id,
        ok: false,
        error: { name: err.name, message: err.message },
      });
    }
  };

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleParentMessage,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(
          "[dialogHost] addHandlerAsync(DialogParentMessageReceived) failed:",
          result.error.message,
        );
      }
    },
  );

  // --- Step 4: Boot the runtime (ensureReady), pushing progress events ---
  //
  // We call ensureReady() eagerly so Pyodide is warm before the first request.
  // Progress messages are forwarded to the parent as BridgeEvent{progress}.

  setStatus("Loading Pyodide and pyspark-connect-web…");

  try {
    await bridge.ensureReady();
  } catch (bootErr) {
    const err = bootErr instanceof Error ? bootErr : new Error(String(bootErr));
    const msg = `Spark engine failed to start: ${err.message}`;
    setStatus(msg, /* isError */ true);
    pushEvent({
      kind: "evt",
      event: "log",
      payload: { level: "error", message: msg },
    });
    return;
  }

  // --- Step 5: Announce readiness to the parent ---
  const readyStatus = bridge.status();
  setStatus("Spark engine ready.");
  pushEvent({
    kind: "evt",
    event: "ready",
    payload: readyStatus,
  });
}

// ---------------------------------------------------------------------------
// Entry: wait for Office.js to initialise, then run the boot sequence.
// ---------------------------------------------------------------------------

Office.onReady(() => {
  main().catch((err: unknown) => {
    const message = err instanceof Error ? err.message : String(err);
    setStatus(`Unexpected error: ${message}`, /* isError */ true);
    console.error("[dialogHost] Unhandled error in main():", err);
  });
});
