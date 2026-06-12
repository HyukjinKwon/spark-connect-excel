// SPDX-License-Identifier: Apache-2.0
//
// query-flow.spec.ts — Spark/Excel-dependent end-to-end matrix.
//
// These tests verify the full connect → runSQL → range → chart flow against a
// real Spark Connect server and a real Excel host (sideloaded add-in).
//
// WHY THESE ARE DEFERRED
// ----------------------
// Running against Excel requires:
//   1. A sideloaded add-in (Windows: registry; Mac: ~/Library/...; Web: admin).
//   2. A running Spark Connect + Envoy stack (docker compose up in deploy/).
//   3. Playwright's Office.js interop or a custom bridge harness — there is no
//      standard Playwright support for Excel add-in task-pane interaction.
//
// None of these are available in GitHub Actions CI without a self-hosted runner
// with Excel installed or a cloud Office testing service. The COI gate
// (coi.spec.ts) is the automatable core; these tests are the staged verification
// that a human maintainer (or a future CI enhancement) runs manually.
//
// HOW TO ENABLE
// -------------
// Set E2E_REQUIRE_STACK=true in your environment before running:
//
//   E2E_REQUIRE_STACK=true npx playwright test query-flow.spec.ts
//
// Without that variable every test is skipped. This mirrors the pattern used in
// pyspark-connect-web's own e2e scaffold.
//
// MANUAL VERIFICATION CHECKLIST (run against the deploy/ stack)
// -------------------------------------------------------------
//   1. docker compose -f deploy/compose.yaml up    (wait ~60s for Spark healthy)
//   2. npm run build && npm run preview             (or the deploy static host)
//   3. Sideload manifest.xml into Excel
//      (see scripts/sideload.md for instructions per platform)
//   4. Open the task pane → click "Spark SQL" button
//   5. In the Connection form:
//        host = localhost, port = 8081, TLS = off
//      Click Connect — status should show "connected"
//   6. Enter: SELECT range(1, 11) AS n  — row cap 10
//      Click Run → expect 10 rows landed in Sheet1 starting at A1
//      Header row: "n"  |  Body rows: 1 … 10
//   7. Click Refresh → same range re-populated (idempotent)
//   8. Click Insert Chart → column chart appears to the right of the data
//   9. Edit the SQL to: SELECT range(1, 6) AS n
//      Click Run → 5 rows; previously truncated range shrinks correctly
//  10. Close Excel (save the workbook), reopen it, click Refresh
//      → query binding persists; result is re-fetched from Spark

import { test } from "@playwright/test";

// Guard: skip every test in this file unless the caller explicitly opts in.
const requireStack = Boolean(process.env["E2E_REQUIRE_STACK"]);

// ---------------------------------------------------------------------------
// Helper: annotate deferred tests with a clear TODO
// ---------------------------------------------------------------------------

/**
 * Declare a deferred test — skipped in CI, fixme-flagged for maintainers.
 * The body is a no-op stub; real assertions go in the TODO comment.
 */
function deferredTest(title: string, todo: string): void {
  // test.fixme marks it as "known-failing / not yet implemented" in the report.
  test.fixme(
    title,
     
    async () => {
      void todo; // reference so linter doesn't complain
      // TODO: implement once the Excel + Spark test harness is available.
      // See the manual checklist in this file's header comment.
    },
  );
}

// ---------------------------------------------------------------------------
// Suite — only registered when E2E_REQUIRE_STACK is set
// ---------------------------------------------------------------------------

test.describe("Spark/Excel full flow (requires E2E_REQUIRE_STACK=true)", () => {
  test.skip(!requireStack, "Set E2E_REQUIRE_STACK=true to run this suite");

  test("task pane loads and reports connected status after connect()", async ({
    page,
  }) => {
    // TODO: Navigate to task pane via the sideloaded add-in mechanism.
    // The task pane URL is https://localhost:3000/taskpane/taskpane.html when
    // using `npm run preview`. In a real Excel sideload the URL is injected by
    // manifest.xml SourceLocation.
    //
    // Steps:
    //   1. page.goto("/taskpane/taskpane.html")
    //   2. Fill in host=localhost, port=8081, TLS=off
    //   3. Click "Connect"
    //   4. await page.locator('[data-testid="status-connected"]').toBeVisible()
    //
    // Requires: Office.js mock or a custom harness that stubs Office.context.
    void page;
    throw new Error("Not implemented — see TODO above");
  });

  deferredTest(
    "runSQL lands header row + typed rows in Sheet1",
    `
    TODO:
      1. Connect (as above)
      2. Enter SQL: SELECT id, name, revenue FROM demo.customers LIMIT 5
      3. Click Run
      4. Assert: Sheet1!A1 = "id", B1 = "name", C1 = "revenue"
      5. Assert: A2 is a number (bigint column)
      6. Assert: rowCount label shows "5 rows"
    `,
  );

  deferredTest(
    "truncation banner appears when rowCap is exceeded",
    `
    TODO:
      1. Connect
      2. Enter SQL: SELECT * FROM range(1, 20001) AS t (20000 rows)
      3. Set row cap to 10000
      4. Click Run
      5. Assert: truncation banner is visible in the sheet (amber background)
      6. Assert: rowCount = 10000 (not 20000)
      7. Assert: result.truncated = true
    `,
  );

  deferredTest(
    "Refresh re-runs the query and updates the same range",
    `
    TODO:
      1. Run a query that lands rows in A1:B11
      2. Click Refresh
      3. Assert: range A1:B11 is rewritten (same header, same or updated data)
      4. Verify that the queryId binding persists in document settings
    `,
  );

  deferredTest(
    "Insert Chart produces a native Excel chart from the result range",
    `
    TODO:
      1. Run: SELECT category, SUM(revenue) AS total FROM sales GROUP BY category
      2. Click "Insert Chart"
      3. Assert: a chart object exists on the sheet
         (Excel.run → worksheet.charts.items.length > 0)
      4. Assert: the chart is a columnClustered type (inferred from 1-categorical + 1-numeric)
      5. Assert: the chart title contains column names from inferChartType
    `,
  );

  deferredTest(
    "Time-series query → line chart inferred",
    `
    TODO:
      1. Run: SELECT event_date, count(*) AS events FROM logs GROUP BY event_date ORDER BY 1
      2. Click "Insert Chart"
      3. Assert: chart type is Excel.ChartType.line (temporal + numeric → line rule)
    `,
  );

  deferredTest(
    "Token persists in OfficeRuntime.storage, never in document settings",
    `
    TODO:
      1. Connect with a bearer token
      2. Assert: Office.context.document.settings has no key containing the token value
      3. Assert: OfficeRuntime.storage has the token (or verify via saveToken / loadToken)
      4. Save + reopen the workbook
      5. Assert: the token is not visible in the workbook XML (unzip .xlsx, check settings.xml)
    `,
  );

  deferredTest(
    "Query binding survives workbook save + reopen",
    `
    TODO:
      1. Run a query; note the queryId from binding.ts
      2. Save and reopen the workbook (or call Office.context.document.settings.reload())
      3. Assert: loadQueryBindings() returns the same record
      4. Click Refresh → same range is rewritten from Spark
    `,
  );

  deferredTest(
    "cancel() interrupts an in-flight query",
    `
    TODO:
      1. Issue a long-running query (e.g. SELECT sleep(30))
      2. Before it completes, click Cancel
      3. Assert: the pending runSQL promise rejects (or resolves with 0 rows)
      4. Assert: status() shows connected=false after cancel (SparkBridgeHost resets)
    `,
  );
});
