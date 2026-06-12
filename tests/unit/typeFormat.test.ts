// SPDX-License-Identifier: Apache-2.0
//
// typeFormat.test.ts — Lane F unit tests
//
// Tests for the pure helpers in src/excel/typeFormat.ts.
// These do NOT require Office.js; they run with plain Vitest.

import { describe, it, expect } from "vitest";
import {
  numberFormatFor,
  coerceCellValue,
  isoToExcelSerial,
} from "../../src/excel/typeFormat.js";

// ---------------------------------------------------------------------------
// numberFormatFor
// ---------------------------------------------------------------------------

describe("numberFormatFor", () => {
  describe("integer types", () => {
    it.each([
      ["bigint", "0"],
      ["int", "0"],
      ["integer", "0"],
      ["smallint", "0"],
      ["tinyint", "0"],
      ["long", "0"],
      ["short", "0"],
      ["byte", "0"],
    ])("maps %s → %s", (sparkType, expected) => {
      expect(numberFormatFor(sparkType)).toBe(expected);
    });

    it("is case-insensitive", () => {
      expect(numberFormatFor("BIGINT")).toBe("0");
      expect(numberFormatFor("Int")).toBe("0");
    });
  });

  describe("floating-point types", () => {
    it.each([
      ["double", "0.00"],
      ["float", "0.00"],
      ["real", "0.00"],
    ])("maps %s → %s", (sparkType, expected) => {
      expect(numberFormatFor(sparkType)).toBe(expected);
    });
  });

  describe("decimal types", () => {
    it("maps decimal(10,2) → 0.00", () => {
      expect(numberFormatFor("decimal(10,2)")).toBe("0.00");
    });

    it("maps decimal(18,0) → 0 (integer scale)", () => {
      expect(numberFormatFor("decimal(18,0)")).toBe("0");
    });

    it("maps decimal(38,6) → 0.000000", () => {
      expect(numberFormatFor("decimal(38,6)")).toBe("0.000000");
    });

    it("maps bare decimal (no precision/scale) → 0.00 (default scale 2)", () => {
      expect(numberFormatFor("decimal")).toBe("0.00");
    });

    it("maps decimal(5, 3) with a space after comma → 0.000", () => {
      expect(numberFormatFor("decimal(5, 3)")).toBe("0.000");
    });

    it("maps DECIMAL(10,2) case-insensitively", () => {
      expect(numberFormatFor("DECIMAL(10,2)")).toBe("0.00");
    });
  });

  describe("numeric types (ANSI alias for decimal)", () => {
    it("maps numeric → 0.00 (bare, default scale 2)", () => {
      expect(numberFormatFor("numeric")).toBe("0.00");
    });

    it("maps numeric(10,2) → 0.00", () => {
      expect(numberFormatFor("numeric(10,2)")).toBe("0.00");
    });

    it("maps numeric(18,0) → 0 (integer scale)", () => {
      expect(numberFormatFor("numeric(18,0)")).toBe("0");
    });

    it("maps numeric(38,6) → 0.000000", () => {
      expect(numberFormatFor("numeric(38,6)")).toBe("0.000000");
    });

    it("maps NUMERIC(10,2) case-insensitively", () => {
      expect(numberFormatFor("NUMERIC(10,2)")).toBe("0.00");
    });
  });

  describe("date / timestamp types", () => {
    it("maps date → yyyy-mm-dd", () => {
      expect(numberFormatFor("date")).toBe("yyyy-mm-dd");
    });

    it("maps timestamp → yyyy-mm-dd hh:mm:ss", () => {
      expect(numberFormatFor("timestamp")).toBe("yyyy-mm-dd hh:mm:ss");
    });

    it("maps timestamp_ntz → yyyy-mm-dd hh:mm:ss", () => {
      expect(numberFormatFor("timestamp_ntz")).toBe("yyyy-mm-dd hh:mm:ss");
    });

    it("maps timestamp_ltz → yyyy-mm-dd hh:mm:ss", () => {
      expect(numberFormatFor("timestamp_ltz")).toBe("yyyy-mm-dd hh:mm:ss");
    });
  });

  describe("text / complex types → null (General format)", () => {
    it.each([
      ["string"],
      ["boolean"],
      ["binary"],
      ["void"],
      ["array<int>"],
      ["map<string,int>"],
      ["struct<a:int,b:string>"],
      ["unknown_future_type"],
    ])("maps %s → null", (sparkType) => {
      expect(numberFormatFor(sparkType)).toBeNull();
    });
  });

  describe("whitespace trimming", () => {
    it("handles leading/trailing whitespace", () => {
      expect(numberFormatFor("  bigint  ")).toBe("0");
      expect(numberFormatFor(" date ")).toBe("yyyy-mm-dd");
    });
  });
});

// ---------------------------------------------------------------------------
// coerceCellValue
// ---------------------------------------------------------------------------

describe("coerceCellValue", () => {
  describe("null / undefined passthrough", () => {
    it("returns null for null", () => {
      expect(coerceCellValue(null, "string")).toBeNull();
    });

    it("returns null for undefined", () => {
      expect(coerceCellValue(undefined, "bigint")).toBeNull();
    });
  });

  describe("boolean passthrough", () => {
    it("returns true for boolean true", () => {
      expect(coerceCellValue(true, "boolean")).toBe(true);
    });

    it("returns false for boolean false", () => {
      expect(coerceCellValue(false, "boolean")).toBe(false);
    });
  });

  describe("numeric types", () => {
    it("passes through integer numbers", () => {
      expect(coerceCellValue(42, "bigint")).toBe(42);
    });

    it("passes through float numbers", () => {
      expect(coerceCellValue(3.14, "double")).toBe(3.14);
    });

    it("passes through decimal numbers", () => {
      expect(coerceCellValue(123.456, "decimal(10,3)")).toBe(123.456);
    });

    it("passes through negative numbers", () => {
      expect(coerceCellValue(-99, "int")).toBe(-99);
    });

    it("passes through zero", () => {
      expect(coerceCellValue(0, "bigint")).toBe(0);
    });
  });

  describe("date type — ISO string → Excel serial", () => {
    // 1970-01-01 is the Unix epoch = Excel serial 25569.
    it("converts 1970-01-01 to Excel serial 25569", () => {
      expect(coerceCellValue("1970-01-01", "date")).toBe(25569);
    });

    // 2024-03-15 sanity-check: days since 1970-01-01 + 25569.
    // 2024-03-15T00:00:00Z → Unix ms = 1710460800000
    // days = 1710460800000 / 86400000 = 19797
    // serial = 19797 + 25569 = 45366
    it("converts 2024-03-15 to correct Excel serial", () => {
      const result = coerceCellValue("2024-03-15", "date");
      expect(result).toBe(45366);
    });

    it("returns string fallback for invalid date", () => {
      expect(coerceCellValue("not-a-date", "date")).toBe("not-a-date");
    });

    it("returns null for null date", () => {
      expect(coerceCellValue(null, "date")).toBeNull();
    });
  });

  describe("timestamp type — ISO string → Excel serial (with fraction)", () => {
    // 1970-01-01T00:00:00Z → serial 25569.0
    it("converts epoch timestamp to 25569.0", () => {
      expect(coerceCellValue("1970-01-01T00:00:00Z", "timestamp")).toBe(25569);
    });

    // 1970-01-01T12:00:00Z → serial 25569.5
    it("converts noon epoch timestamp to 25569.5", () => {
      expect(coerceCellValue("1970-01-01T12:00:00Z", "timestamp")).toBe(25569.5);
    });

    // timestamp_ntz — treated the same way (appended 'Z' internally).
    it("handles timestamp_ntz with no timezone info", () => {
      expect(coerceCellValue("1970-01-01T00:00:00", "timestamp_ntz")).toBe(25569);
    });

    it("returns string fallback for invalid timestamp", () => {
      expect(coerceCellValue("bad-ts", "timestamp")).toBe("bad-ts");
    });

    it("returns null for null timestamp", () => {
      expect(coerceCellValue(null, "timestamp")).toBeNull();
    });
  });

  describe("string type", () => {
    it("returns the string unchanged", () => {
      expect(coerceCellValue("hello", "string")).toBe("hello");
    });

    it("converts a non-string value to string", () => {
      // Unlikely from SparkResult but should be safe.
      expect(coerceCellValue(42, "string")).toBe(42);
    });
  });

  describe("complex / unknown types", () => {
    it("converts objects to string via String()", () => {
      // Complex types come through as JSON strings from the Python marshaller;
      // if a raw object arrives we stringify it.
      expect(coerceCellValue("[1,2,3]", "array<int>")).toBe("[1,2,3]");
    });
  });
});

// ---------------------------------------------------------------------------
// isoToExcelSerial (exported internal helper — directly testable)
// ---------------------------------------------------------------------------

describe("isoToExcelSerial", () => {
  it("1899-12-30 is Excel serial 0 (epoch)", () => {
    // 1899-12-30T00:00:00Z in Unix ms:
    // days before 1970-01-01: 25569 days → ms = -25569 * 86400000
    // serial = -25569 / 1 + 25569 = 0
    expect(isoToExcelSerial("1899-12-30", true)).toBe(0);
  });

  it("1900-01-01 is Excel serial 2 (Lotus leap-day compat: 1=1900-01-00, 2=1900-01-01)", () => {
    // Excel serial 2 = 1900-01-01 (due to the Lotus 1-2-3 bug)
    // Calculation: days from 1899-12-30 to 1900-01-01 = 2
    expect(isoToExcelSerial("1900-01-01", true)).toBe(2);
  });

  it("returns the raw string for unparseable input", () => {
    expect(isoToExcelSerial("not-a-date", false)).toBe("not-a-date");
  });

  it("handles datetime with explicit Z timezone", () => {
    const result = isoToExcelSerial("1970-01-01T06:00:00Z", false);
    // 6 hours = 0.25 day → serial = 25569.25
    expect(result).toBe(25569.25);
  });

  it("handles datetime with no timezone (treated as UTC)", () => {
    const result = isoToExcelSerial("1970-01-01T06:00:00", false);
    expect(result).toBe(25569.25);
  });

  it("handles datetime with positive offset +05:30", () => {
    // "1970-01-01T05:30:00+05:30" is UTC 00:00:00 → serial 25569
    const result = isoToExcelSerial("1970-01-01T05:30:00+05:30", false);
    expect(result).toBe(25569);
  });
});
