// SPDX-License-Identifier: Apache-2.0
//
// coi.ts — Cross-origin isolation gate (Lane B).
//
// Responsibilities:
//   1. `ensureCrossOriginIsolated()` — verify or trigger COI via the
//      pyspark-connect-web coi-serviceworker.js shim (belt-and-suspenders; the
//      dialog's actual isolation comes from COOP/COEP headers the dev server and
//      production static host send, per DECISIONS.md #1/#2).
//   2. `isolationDiagnostics()` — cheap, synchronous status snapshot for the
//      parent progress feed and the findings doc.
//
// NOTE: This file runs inside the Office Dialog window, NOT the task pane.
//       The service-worker scope is the dialog's own origin (same-origin).

/**
 * Diagnostic snapshot — cheap, synchronous, safe to call at any time.
 *
 * @returns
 *   `crossOriginIsolated` — `self.crossOriginIsolated`; true means
 *     SharedArrayBuffer/Atomics are available.
 *   `sharedArrayBuffer` — whether `SharedArrayBuffer` is constructible right now.
 */
export function isolationDiagnostics(): {
  crossOriginIsolated: boolean;
  sharedArrayBuffer: boolean;
} {
  return {
    crossOriginIsolated: typeof self !== "undefined" ? self.crossOriginIsolated : false,
    sharedArrayBuffer: typeof SharedArrayBuffer !== "undefined",
  };
}

/**
 * Ensure the dialog window is cross-origin isolated so that SharedArrayBuffer
 * and Atomics are available (required by the pyspark-connect-web SAB bridge).
 *
 * Strategy (belt-and-suspenders):
 *   1. If `self.crossOriginIsolated` is already true, return true immediately.
 *      This is the normal path when the dev server / deploy host sends the
 *      COOP/COEP headers (DECISIONS.md #1/#2).
 *   2. Otherwise, attempt to register the `coi-serviceworker.js` shim, which
 *      intercepts fetch responses and re-emits them with COOP/COEP so that the
 *      page becomes isolated after a one-time reload.  The shim itself guards
 *      against reload loops via sessionStorage (`pcwCoiReloaded`), so we do NOT
 *      add our own guard here.
 *   3. Return `self.crossOriginIsolated` after registration.  If we just
 *      registered the SW for the first time the page will reload; if the
 *      reload already happened and isolation still is false, this returns false
 *      and the caller should surface an error.
 *
 * @returns true if the window is cross-origin isolated; false otherwise.
 */
export async function ensureCrossOriginIsolated(): Promise<boolean> {
  // Fast path: already isolated (COOP/COEP headers came from the server).
  if (self.crossOriginIsolated) {
    return true;
  }

  // Only proceed if the browser exposes a service-worker container.
  if (typeof navigator === "undefined" || !("serviceWorker" in navigator)) {
    console.warn("[coi] Service workers not available; cannot enforce COI.");
    return self.crossOriginIsolated;
  }

  // The coi-serviceworker.js shim has its own sessionStorage loop-guard
  // (`pcwCoiReloaded`).  We simply register it; if it decides a reload is
  // needed it will trigger one and the page will restart.
  try {
    await navigator.serviceWorker.register("/vendor/coi-serviceworker.js", {
      scope: "./",
    });
  } catch (err) {
    console.error("[coi] Failed to register coi-serviceworker:", err);
  }

  // Return the current state.  If the SW just registered and a reload is
  // imminent, this will still be false — but the page is about to restart and
  // ensureCrossOriginIsolated will return true on the reloaded page.
  return self.crossOriginIsolated;
}
