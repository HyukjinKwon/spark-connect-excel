// SPDX-License-Identifier: Apache-2.0
//
// chartView.ts - pure SVG chart renderer for SparkResult (web demo only).
//
// Renders an inline SVG bar chart (categorical X) or line chart (temporal X)
// from a SparkResult with no external dependencies. Kept free of any runtime /
// Office dependency so it is unit-testable under jsdom.

import type { SparkResult } from "../seam";

// ---------------------------------------------------------------------------
// Spark type classifiers (mirror of src/excel/chart.ts - exported for tests)
// ---------------------------------------------------------------------------

/** Returns true for numeric Spark SQL type names. */
export function isNumeric(type: string): boolean {
  const t = type.toLowerCase();
  return (
    t === "bigint" ||
    t === "int" ||
    t === "integer" ||
    t === "smallint" ||
    t === "tinyint" ||
    t === "double" ||
    t === "float" ||
    t === "real" ||
    t.startsWith("decimal") ||
    t.startsWith("numeric")
  );
}

/** Returns true for temporal Spark SQL type names. */
export function isTemporal(type: string): boolean {
  const t = type.toLowerCase();
  return t === "date" || t === "timestamp" || t === "timestamp_ntz" || t === "timestamp_ltz";
}

/** Returns true for categorical (string-like or boolean) Spark SQL type names. */
export function isCategorical(type: string): boolean {
  const t = type.toLowerCase();
  return (
    t === "string" ||
    t === "boolean" ||
    t === "bool" ||
    t.startsWith("char") ||
    t.startsWith("varchar")
  );
}

// ---------------------------------------------------------------------------
// SVG namespace constant
// ---------------------------------------------------------------------------

const SVG_NS = "http://www.w3.org/2000/svg";

// ---------------------------------------------------------------------------
// Chart layout constants
// ---------------------------------------------------------------------------

const VIEW_W = 520;
const VIEW_H = 300;
const PAD_TOP = 36; // title + breathing room
const PAD_RIGHT = 20;
const PAD_BOTTOM = 56; // X labels
const PAD_LEFT = 56; // Y axis labels
const MAX_CATEGORIES = 16;

// Colours matching the demo palette
const ORANGE = "#e25a1c";
const BLUE = "#2272b4";
const INK = "#1f2328";
const MUTED = "#57606a";
const LINE = "#d0d7de";

// ---------------------------------------------------------------------------
// Safe SVG text helper
// ---------------------------------------------------------------------------

/** Escape a string so it is safe to embed in an SVG text element. */
function escapeText(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** Create an SVG element with attributes. */
function svgEl(tag: string, attrs: Record<string, string | number>): SVGElement {
  const el = document.createElementNS(SVG_NS, tag);
  for (const [k, v] of Object.entries(attrs)) {
    el.setAttribute(k, String(v));
  }
  return el;
}

/** Create an SVG text element. */
function svgText(content: string, attrs: Record<string, string | number>): SVGElement {
  const el = svgEl("text", attrs);
  el.textContent = escapeText(content);
  return el;
}

// ---------------------------------------------------------------------------
// Number formatter
// ---------------------------------------------------------------------------

/** Format a number compactly (e.g. 1200 -> "1.2k", 1500000 -> "1.5M"). */
function fmtNum(n: number): string {
  const abs = Math.abs(n);
  if (abs >= 1_000_000) return `${(n / 1_000_000).toFixed(1)}M`;
  if (abs >= 1_000) return `${(n / 1_000).toFixed(1)}k`;
  if (!Number.isInteger(n)) return n.toFixed(2);
  return String(n);
}

/** Truncate a label to at most maxLen chars, appending "..." if cut. */
function truncLabel(s: string, maxLen = 10): string {
  if (s.length <= maxLen) return s;
  return s.slice(0, maxLen - 3) + "...";
}

// ---------------------------------------------------------------------------
// renderChart - main export
// ---------------------------------------------------------------------------

/**
 * Render a SparkResult as an inline SVG chart element.
 *
 * Returns null when the result is not chartable:
 *   - no rows, OR
 *   - no usable numeric column.
 *
 * Heuristic (mirrors chart.ts inferChartType):
 *   - temporal X + numeric Y   -> line chart (blue line + dots)
 *   - categorical X + numeric Y -> bar chart (orange bars)
 *   - >=2 numeric, no category  -> bar chart of first numeric by row index
 *   - no numeric column          -> null
 */
export function renderChart(result: SparkResult): HTMLElement | null {
  if (result.rows.length === 0) return null;

  const schema = result.schema;
  const numericIdx = schema.findIndex((c) => isNumeric(c.type));
  if (numericIdx === -1) return null;

  const temporalIdx = schema.findIndex((c) => isTemporal(c.type));
  const categoricalIdx = schema.findIndex((c) => isCategorical(c.type));

  let xIdx: number;
  let chartType: "bar" | "line";

  if (temporalIdx !== -1) {
    // temporal + numeric -> line
    xIdx = temporalIdx;
    chartType = "line";
  } else if (categoricalIdx !== -1) {
    // categorical + numeric -> bar
    xIdx = categoricalIdx;
    chartType = "bar";
  } else {
    // >=2 numeric or only numeric: bar by row index
    xIdx = -1;
    chartType = "bar";
  }

  const yIdx = numericIdx;
  const yCol = schema[yIdx].name;
  const xCol = xIdx >= 0 ? schema[xIdx].name : "index";

  // Extract data points, cap at MAX_CATEGORIES
  const allRows = result.rows;
  const capped = allRows.length > MAX_CATEGORIES;
  const rows = capped ? allRows.slice(0, MAX_CATEGORIES) : allRows;

  const labels: string[] = rows.map((row, i) => {
    if (xIdx < 0) return String(i + 1);
    const v = row[xIdx];
    if (v === null || v === undefined) return "null";
    return String(v);
  });

  const values: (number | null)[] = rows.map((row) => {
    const v = row[yIdx];
    if (v === null || v === undefined) return null;
    const n = Number(v);
    return isNaN(n) ? null : n;
  });

  const title = xIdx >= 0 ? `${yCol} by ${xCol}` : yCol;

  const extraNote = capped ? `+${allRows.length - MAX_CATEGORIES} more` : null;

  // Build the SVG
  const svg = buildSVG(labels, values, title, extraNote, chartType);

  // Wrap in a container div
  const wrapper = document.createElement("div");
  wrapper.className = "demo-chart";
  wrapper.appendChild(svg);
  return wrapper;
}

// ---------------------------------------------------------------------------
// SVG builder
// ---------------------------------------------------------------------------

function buildSVG(
  labels: string[],
  values: (number | null)[],
  title: string,
  extraNote: string | null,
  chartType: "bar" | "line",
): SVGElement {
  const svg = svgEl("svg", {
    viewBox: `0 0 ${VIEW_W} ${VIEW_H}`,
    width: VIEW_W,
    height: VIEW_H,
    "aria-label": escapeText(title),
    role: "img",
  });

  const plotW = VIEW_W - PAD_LEFT - PAD_RIGHT;
  const plotH = VIEW_H - PAD_TOP - PAD_BOTTOM;

  // Compute data range
  const nums = values.filter((v): v is number => v !== null);
  if (nums.length === 0) {
    // Render empty state
    const msg = svgText("No numeric data to display", {
      x: VIEW_W / 2,
      y: VIEW_H / 2,
      "text-anchor": "middle",
      "dominant-baseline": "middle",
      fill: MUTED,
      "font-size": "13",
      "font-family": "inherit",
    });
    svg.appendChild(msg);
    return svg;
  }

  const rawMin = Math.min(...nums);
  const rawMax = Math.max(...nums);
  // For bar charts, baseline is at 0 (or min if all negative)
  const yMin = chartType === "bar" ? Math.min(0, rawMin) : rawMin * 0.95;
  const yMax = rawMax === yMin ? yMin + 1 : rawMax + (rawMax - yMin) * 0.1;

  const yRange = yMax - yMin;

  function toY(v: number): number {
    return PAD_TOP + plotH - ((v - yMin) / yRange) * plotH;
  }

  // -- Background
  svg.appendChild(
    svgEl("rect", {
      x: 0,
      y: 0,
      width: VIEW_W,
      height: VIEW_H,
      fill: "#ffffff",
      rx: 0,
    }),
  );

  // -- Title
  svg.appendChild(
    svgText(title, {
      x: VIEW_W / 2,
      y: 16,
      "text-anchor": "middle",
      "dominant-baseline": "middle",
      fill: INK,
      "font-size": "13",
      "font-weight": "600",
      "font-family": "inherit",
    }),
  );

  // -- Y axis grid lines + labels (5 ticks)
  const TICK_COUNT = 5;
  for (let i = 0; i <= TICK_COUNT; i++) {
    const v = yMin + (yRange * i) / TICK_COUNT;
    const py = toY(v);
    // Grid line
    svg.appendChild(
      svgEl("line", {
        x1: PAD_LEFT,
        y1: py,
        x2: PAD_LEFT + plotW,
        y2: py,
        stroke: i === 0 ? LINE : "#ebebeb",
        "stroke-width": i === 0 ? 1 : 0.5,
        "stroke-dasharray": i === 0 ? "none" : "3,3",
      }),
    );
    // Y label
    svg.appendChild(
      svgText(fmtNum(v), {
        x: PAD_LEFT - 6,
        y: py,
        "text-anchor": "end",
        "dominant-baseline": "middle",
        fill: MUTED,
        "font-size": "10",
        "font-family": "inherit",
      }),
    );
  }

  // -- Baseline axis (y=0 line for bar charts if 0 is in range)
  const baselineY = toY(0);
  if (baselineY >= PAD_TOP && baselineY <= PAD_TOP + plotH) {
    svg.appendChild(
      svgEl("line", {
        x1: PAD_LEFT,
        y1: baselineY,
        x2: PAD_LEFT + plotW,
        y2: baselineY,
        stroke: LINE,
        "stroke-width": "1.5",
      }),
    );
  }

  // -- Y axis vertical line
  svg.appendChild(
    svgEl("line", {
      x1: PAD_LEFT,
      y1: PAD_TOP,
      x2: PAD_LEFT,
      y2: PAD_TOP + plotH,
      stroke: LINE,
      "stroke-width": "1",
    }),
  );

  const n = labels.length;

  if (chartType === "bar") {
    const barGap = 0.15;
    const slotW = plotW / n;
    const barW = slotW * (1 - barGap * 2);

    for (let i = 0; i < n; i++) {
      const v = values[i];
      const cx = PAD_LEFT + slotW * i + slotW / 2;

      // X label
      svg.appendChild(
        svgText(truncLabel(labels[i], 9), {
          x: cx,
          y: PAD_TOP + plotH + 14,
          "text-anchor": "middle",
          "dominant-baseline": "middle",
          fill: MUTED,
          "font-size": "10",
          "font-family": "inherit",
        }),
      );

      if (v === null) continue;

      const barBase = toY(Math.max(0, yMin < 0 ? 0 : yMin));
      const barTop = toY(v);
      const barH = Math.abs(barBase - barTop);
      const barY = Math.min(barBase, barTop);

      // Bar rect
      svg.appendChild(
        svgEl("rect", {
          x: cx - barW / 2,
          y: barY,
          width: barW,
          height: Math.max(barH, 1),
          fill: ORANGE,
          rx: "2",
        }),
      );

      // Value label on bar (above or below depending on sign)
      const labelY = v >= 0 ? barY - 4 : barY + barH + 12;
      svg.appendChild(
        svgText(fmtNum(v), {
          x: cx,
          y: labelY,
          "text-anchor": "middle",
          "dominant-baseline": v >= 0 ? "auto" : "auto",
          fill: INK,
          "font-size": "9.5",
          "font-weight": "600",
          "font-family": "inherit",
        }),
      );
    }
  } else {
    // Line chart
    const stepX = plotW / Math.max(n - 1, 1);

    // Build path
    const pathParts: string[] = [];
    const points: { x: number; y: number; v: number }[] = [];

    for (let i = 0; i < n; i++) {
      const v = values[i];
      if (v === null) continue;
      const px = PAD_LEFT + (n === 1 ? plotW / 2 : i * stepX);
      const py = toY(v);
      points.push({ x: px, y: py, v });
      pathParts.push(pathParts.length === 0 ? `M ${px} ${py}` : `L ${px} ${py}`);
    }

    if (pathParts.length > 0) {
      // Fill area under line
      const firstPt = points[0];
      const lastPt = points[points.length - 1];
      const areaPath =
        pathParts.join(" ") + ` L ${lastPt.x} ${baselineY} L ${firstPt.x} ${baselineY} Z`;

      svg.appendChild(
        svgEl("path", {
          d: areaPath,
          fill: BLUE,
          opacity: "0.08",
          stroke: "none",
        }),
      );

      // Line
      svg.appendChild(
        svgEl("path", {
          d: pathParts.join(" "),
          fill: "none",
          stroke: BLUE,
          "stroke-width": "2",
          "stroke-linejoin": "round",
          "stroke-linecap": "round",
        }),
      );

      // Dots + value labels
      for (const pt of points) {
        svg.appendChild(
          svgEl("circle", {
            cx: pt.x,
            cy: pt.y,
            r: "3.5",
            fill: "#ffffff",
            stroke: BLUE,
            "stroke-width": "2",
          }),
        );
        svg.appendChild(
          svgText(fmtNum(pt.v), {
            x: pt.x,
            y: pt.y - 9,
            "text-anchor": "middle",
            fill: BLUE,
            "font-size": "9",
            "font-weight": "600",
            "font-family": "inherit",
          }),
        );
      }
    }

    // X labels
    for (let i = 0; i < n; i++) {
      const px = PAD_LEFT + (n === 1 ? plotW / 2 : i * stepX);
      svg.appendChild(
        svgText(truncLabel(labels[i], 10), {
          x: px,
          y: PAD_TOP + plotH + 14,
          "text-anchor": "middle",
          fill: MUTED,
          "font-size": "10",
          "font-family": "inherit",
        }),
      );
    }
  }

  // -- "+N more" note if capped
  if (extraNote) {
    svg.appendChild(
      svgText(extraNote, {
        x: VIEW_W - PAD_RIGHT,
        y: PAD_TOP + plotH + 38,
        "text-anchor": "end",
        fill: MUTED,
        "font-size": "10",
        "font-style": "italic",
        "font-family": "inherit",
      }),
    );
  }

  return svg;
}
