import { describe, expect, it } from "vitest";
import {
  cellRef,
  columnLetter,
  qualify,
  rangeRef,
  sanitizeExcelName,
  sanitizeSheetName,
  splitAddress,
  uniqueName,
} from "../src/core/a1";
import { niceFloorBelow } from "../src/core/numeric";

describe("A1 helpers", () => {
  it("column letters", () => {
    expect([0, 1, 25, 26, 27, 51, 52, 701, 702, 16383].map(columnLetter)).toEqual([
      "A",
      "B",
      "Z",
      "AA",
      "AB",
      "AZ",
      "BA",
      "ZZ",
      "AAA",
      "XFD",
    ]);
  });

  it("cell and range references", () => {
    expect(cellRef(0, 0)).toBe("A1");
    expect(rangeRef(1, 3, 8, 2)).toBe("D2:E9");
    expect(rangeRef(4, 4, 1, 1)).toBe("E5");
  });

  it("splits and qualifies sheet addresses", () => {
    expect(splitAddress("Sheet1!A1:B4")).toEqual({ sheet: "Sheet1", local: "A1:B4" });
    expect(splitAddress("'Bob''s Data'!C3")).toEqual({ sheet: "Bob's Data", local: "C3" });
    expect(splitAddress("A1")).toEqual({ sheet: null, local: "A1" });
    expect(qualify("Bob's Data", "C3")).toBe("'Bob''s Data'!C3");
    expect(qualify("Sheet1", "C3")).toBe("Sheet1!C3");
  });
});

describe("Excel names", () => {
  it("sanitizes table names", () => {
    expect(sanitizeExcelName("CDF Make Component")).toBe("CDF_Make_Component");
    expect(sanitizeExcelName("2024 plan")).toBe("_2024_plan");
    expect(sanitizeExcelName("CDF1")).toBe("CDF1_");
    expect(sanitizeExcelName("R1C1")).toBe("R1C1_");
    expect(sanitizeExcelName("c")).toBe("c_");
    expect(sanitizeExcelName("  ")).toBe("CDF_Table");
    expect(sanitizeExcelName("Coût/€")).toBe("Co_t");
    expect(sanitizeExcelName("CDF Unsorted (percent text)")).toBe("CDF_Unsorted_percent_text");
  });

  it("makes names unique", () => {
    expect(uniqueName("CDF_A", ["cdf_a", "CDF_A_2"])).toBe("CDF_A_3");
    expect(uniqueName("CDF_B", ["CDF_A"])).toBe("CDF_B");
  });

  it("sanitizes sheet names", () => {
    expect(sanitizeSheetName("CDF: Output/2024?")).toBe("CDF  Output 2024");
    expect(sanitizeSheetName("")).toBe("CDF Output");
    expect(sanitizeSheetName("x".repeat(40))).toHaveLength(31);
  });
});

describe("niceFloorBelow", () => {
  it("returns a round number strictly below the minimum", () => {
    expect(niceFloorBelow(-20, 50)).toBe(-30);
    expect(niceFloorBelow(-5, -5)).toBe(-6);
    expect(niceFloorBelow(-0.3, 0.5)).toBe(-0.4);
    expect(niceFloorBelow(-100, -100)).toBeLessThan(-100);
  });
});
