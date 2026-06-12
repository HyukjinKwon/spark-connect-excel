// SPDX-License-Identifier: Apache-2.0
//
// chart.test.ts — Unit tests for inferChartType (pure, no Office.js).
//
// NOTE: Excel.ChartType is an ambient TypeScript enum from @types/office-js.
// Its numeric values (Excel.ChartType.Line etc.) do NOT exist at runtime in
// jsdom — office-js is never loaded. Therefore these tests import and assert
// against the ChartKind string-literal union returned by inferChartType, NOT
// against Excel.ChartType values. toExcelChartType() is NOT tested here because
// it requires a live Office host; it is exercised by Lane J's e2e tests.

import { describe, it, expect } from "vitest";
import { inferChartType } from "../../src/excel/chart.js";
import type { ChartKind } from "../../src/excel/chart.js";
import type { ColumnMeta } from "../../src/seam.js";

// ---------------------------------------------------------------------------
// Helpers for concise column construction
// ---------------------------------------------------------------------------

function col(name: string, type: string): ColumnMeta {
  return { name, type };
}

// ---------------------------------------------------------------------------
// Rule 1: temporal + ≥1 numeric → "line"
// ---------------------------------------------------------------------------

describe("inferChartType — Rule 1: temporal + numeric → line", () => {
  it("date + one numeric → line", () => {
    const schema: ColumnMeta[] = [col("event_date", "date"), col("revenue", "double")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("line");
  });

  it("timestamp + one numeric → line", () => {
    const schema: ColumnMeta[] = [col("ts", "timestamp"), col("count", "bigint")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("line");
  });

  it("timestamp_ntz + multiple numerics → line (temporal wins)", () => {
    const schema: ColumnMeta[] = [
      col("created_at", "timestamp_ntz"),
      col("sales", "decimal(18,2)"),
      col("units", "int"),
    ];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("line");
  });

  it("temporal column alongside a categorical still → line", () => {
    // Temporal + numeric takes priority over categorical + numeric
    const schema: ColumnMeta[] = [
      col("month", "date"),
      col("region", "string"),
      col("revenue", "double"),
    ];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("line");
  });
});

// ---------------------------------------------------------------------------
// Rule 2: exactly 1 categorical + 1 numeric → "columnClustered"
// ---------------------------------------------------------------------------

describe("inferChartType — Rule 2: one categorical + one numeric → columnClustered", () => {
  it("string + bigint → columnClustered", () => {
    const schema: ColumnMeta[] = [col("product", "string"), col("sales", "bigint")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });

  it("boolean + double → columnClustered", () => {
    const schema: ColumnMeta[] = [col("is_active", "boolean"), col("avg_value", "double")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });

  it("varchar + int → columnClustered", () => {
    const schema: ColumnMeta[] = [col("category", "varchar(64)"), col("count", "int")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });

  it("char + decimal → columnClustered", () => {
    const schema: ColumnMeta[] = [col("code", "char(3)"), col("amount", "decimal(10,2)")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });
});

// ---------------------------------------------------------------------------
// Rule 3: 1 categorical + multiple numerics → "columnClustered"
// ---------------------------------------------------------------------------

describe("inferChartType — Rule 3: one categorical + multiple numerics → columnClustered", () => {
  it("string + two doubles → columnClustered", () => {
    const schema: ColumnMeta[] = [
      col("region", "string"),
      col("revenue", "double"),
      col("cost", "double"),
    ];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });

  it("string + three numeric types → columnClustered", () => {
    const schema: ColumnMeta[] = [
      col("segment", "string"),
      col("q1", "float"),
      col("q2", "int"),
      col("q3", "bigint"),
    ];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });
});

// ---------------------------------------------------------------------------
// Rule 4: ≥2 numerics + no categoricals → "xyScatter"
// ---------------------------------------------------------------------------

describe("inferChartType — Rule 4: two numerics, no category → xyScatter", () => {
  it("two doubles → xyScatter", () => {
    const schema: ColumnMeta[] = [col("x", "double"), col("y", "double")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("xyScatter");
  });

  it("int + float → xyScatter", () => {
    const schema: ColumnMeta[] = [col("age", "int"), col("score", "float")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("xyScatter");
  });

  it("three numerics, no categorical → xyScatter", () => {
    const schema: ColumnMeta[] = [
      col("a", "double"),
      col("b", "double"),
      col("c", "decimal(5,2)"),
    ];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("xyScatter");
  });

  it("bigint + numeric(10,0) → xyScatter", () => {
    const schema: ColumnMeta[] = [col("id", "bigint"), col("val", "numeric(10,0)")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("xyScatter");
  });
});

// ---------------------------------------------------------------------------
// Rule 5 (default): ambiguous or single-column schemas → "columnClustered"
// ---------------------------------------------------------------------------

describe("inferChartType — Rule 5 (default): fallback → columnClustered", () => {
  it("single string column → columnClustered (default)", () => {
    const schema: ColumnMeta[] = [col("label", "string")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });

  it("single numeric column → columnClustered (default, no second series)", () => {
    const schema: ColumnMeta[] = [col("value", "double")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });

  it("empty schema → columnClustered (default)", () => {
    const schema: ColumnMeta[] = [];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });

  it("two categoricals, no numeric → columnClustered (default)", () => {
    const schema: ColumnMeta[] = [col("country", "string"), col("city", "string")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });

  it("mixed unknown types → columnClustered (default)", () => {
    const schema: ColumnMeta[] = [col("data", "binary"), col("extra", "array<string>")];
    const result: ChartKind = inferChartType(schema);
    expect(result).toBe("columnClustered");
  });
});

// ---------------------------------------------------------------------------
// Type-classifier edge cases (via inferChartType behaviour)
// ---------------------------------------------------------------------------

describe("inferChartType — Spark type classifier edge cases", () => {
  it("smallint is treated as numeric", () => {
    const schema: ColumnMeta[] = [col("cat", "string"), col("n", "smallint")];
    expect(inferChartType(schema)).toBe("columnClustered");
  });

  it("tinyint is treated as numeric", () => {
    const schema: ColumnMeta[] = [col("x", "tinyint"), col("y", "tinyint")];
    expect(inferChartType(schema)).toBe("xyScatter");
  });

  it("real is treated as numeric", () => {
    const schema: ColumnMeta[] = [col("price", "real"), col("qty", "int")];
    expect(inferChartType(schema)).toBe("xyScatter");
  });

  it("timestamp_ltz is treated as temporal", () => {
    const schema: ColumnMeta[] = [col("ts", "timestamp_ltz"), col("val", "double")];
    expect(inferChartType(schema)).toBe("line");
  });

  it("bool is treated as categorical", () => {
    // bool (Spark alias for boolean) + numeric → columnClustered
    const schema: ColumnMeta[] = [col("flag", "bool"), col("metric", "double")];
    expect(inferChartType(schema)).toBe("columnClustered");
  });

  it("decimal with precision/scale is numeric", () => {
    const schema: ColumnMeta[] = [col("amount", "decimal(18,4)"), col("tax", "decimal(10,2)")];
    expect(inferChartType(schema)).toBe("xyScatter");
  });

  it("type matching is case-insensitive", () => {
    // Spark returns lowercase type names; guard against mixed-case from seam
    const schema: ColumnMeta[] = [col("ts", "TIMESTAMP"), col("v", "DOUBLE")];
    expect(inferChartType(schema)).toBe("line");
  });
});
