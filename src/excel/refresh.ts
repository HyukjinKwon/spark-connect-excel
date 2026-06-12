// SPDX-License-Identifier: Apache-2.0
//
// refresh.ts — re-run a bound query and rewrite its Excel range (Lane G).
//
// Depends on:
//   • src/excel/binding.ts  — loadQueryBindings (Lane G, this lane)
//   • src/excel/rangeWriter.ts — writeResult / WriteResultInfo (Lane F)
//   • src/seam.ts           — SparkBridge / SparkResult
//   • src/connection/connectionStore.ts — SettingsBackend
//
// DECISIONS #6: the bearer token is NEVER in SavedQuery.  Callers that need
// authentication must ensure the SparkBridge is already connected (via
// bridge.connect()) before calling refreshQuery / refreshAll.

import { type SparkBridge } from "../seam.js";
import {
  type SettingsBackend,
  officeDocumentSettingsBackend,
} from "../connection/connectionStore.js";
import { loadQueryBindings } from "./binding.js";
import { writeResult, type WriteResultInfo } from "./rangeWriter.js";

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Re-run a single saved query and rewrite its Excel range in place.
 *
 * Steps:
 *  1. Look up the SavedQuery by `queryId`.
 *  2. Call `bridge.runSQL(q.sql, q.rowCap)`.
 *  3. Call `writeResult(result, { anchorAddress, sheetName })` (Lane F) to
 *     overwrite exactly the same range that was written originally.
 *  4. Return the WriteResultInfo produced by Lane F.
 *
 * Throws a descriptive `Error` when the `queryId` is not found in the
 * persisted bindings.
 *
 * @param queryId  The ID of the query to refresh (from SavedQuery.queryId).
 * @param bridge   A connected SparkBridge instance.
 * @param backend  Settings backend (default: Office document settings).
 */
export async function refreshQuery(
  queryId: string,
  bridge: SparkBridge,
  backend: SettingsBackend = officeDocumentSettingsBackend(),
): Promise<WriteResultInfo> {
  const queries = loadQueryBindings(backend);
  const q = queries.find((r) => r.queryId === queryId);
  if (q == null) {
    throw new Error(
      `refreshQuery: no saved query found with queryId "${queryId}". ` +
        `Available IDs: [${queries.map((r) => r.queryId).join(", ")}]`,
    );
  }

  const result = await bridge.runSQL(q.sql, q.rowCap);
  const info = await writeResult(result, {
    anchorAddress: q.anchorAddress,
    sheetName: q.sheetName,
  });
  return info;
}

/**
 * Re-run ALL saved queries, collecting per-query success/failure.
 *
 * One query failing does NOT abort the rest — all queries are attempted and
 * errors are captured in the returned array.  This mirrors the UX expectation
 * of a "Refresh All" button: best-effort, report what went wrong.
 *
 * @param bridge   A connected SparkBridge instance.
 * @param backend  Settings backend (default: Office document settings).
 * @returns        An array with one entry per saved query, in load order.
 */
export async function refreshAll(
  bridge: SparkBridge,
  backend: SettingsBackend = officeDocumentSettingsBackend(),
): Promise<{ queryId: string; ok: boolean; error?: string }[]> {
  const queries = loadQueryBindings(backend);

  const results = await Promise.allSettled(
    queries.map((q) => refreshQuery(q.queryId, bridge, backend)),
  );

  return queries.map((q, i) => {
    const outcome = results[i];
    if (outcome === undefined || outcome.status === "rejected") {
      const reason: unknown = outcome?.status === "rejected" ? outcome.reason : undefined;
      const message =
        reason instanceof Error
          ? reason.message
          : typeof reason === "string"
            ? reason
            : "Unknown error";
      return { queryId: q.queryId, ok: false, error: message };
    }
    return { queryId: q.queryId, ok: true };
  });
}
