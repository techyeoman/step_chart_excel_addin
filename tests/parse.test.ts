import { describe, expect, it } from "vitest";
import { detectHeader, parseCell, parseDistribution } from "../src/core/parse";

const errors = (r: ReturnType<typeof parseDistribution>) => r.issues.filter((i) => i.severity === "error");

describe("parseCell", () => {
  it("reads numbers, numeric text, percentages and blanks", () => {
    expect(parseCell(3)).toEqual({ kind: "number", value: 3, percentText: false });
    expect(parseCell(" 2.5 ")).toEqual({ kind: "number", value: 2.5, percentText: false });
    expect(parseCell("40%")).toEqual({ kind: "number", value: 0.4, percentText: true });
    expect(parseCell("")).toEqual({ kind: "blank" });
    expect(parseCell(null)).toEqual({ kind: "blank" });
    expect(parseCell("#N/A")).toMatchObject({ kind: "invalid", excelError: true });
    expect(parseCell("abc")).toMatchObject({ kind: "invalid", excelError: false });
    expect(parseCell(true)).toMatchObject({ kind: "invalid" });
  });
});

describe("parseDistribution", () => {
  it("parses the Make Component example without headers", () => {
    const r = parseDistribution([
      [11, 0.4],
      [32, 0.3],
      [53, 0.3],
    ]);
    expect(r.ok).toBe(true);
    expect(r.hasHeader).toBe(false);
    expect(r.distribution!.outcomes).toEqual([
      { x: 11, p: 0.4 },
      { x: 32, p: 0.3 },
      { x: 53, p: 0.3 },
    ]);
  });

  it("detects a header row automatically and honours the override", () => {
    const values = [
      ["Outcome", "Probability"],
      [15, 0.4],
      [30, 0.3],
      [45, 0.3],
    ];
    expect(detectHeader(values)).toBe(true);
    const r = parseDistribution(values);
    expect(r.ok && r.hasHeader).toBe(true);
    expect(r.headers).toEqual(["Outcome", "Probability"]);
    const forced = parseDistribution(values, { headerMode: "no" });
    expect(forced.ok).toBe(false);
    expect(errors(forced)[0].cell).toBe("A1");
    const numericHeaderless = parseDistribution(values.slice(1), { headerMode: "yes" });
    expect(numericHeaderless.ok).toBe(false); // first data row consumed -> sum 0.6
  });

  it("3. sorts unsorted outcomes and says so", () => {
    const r = parseDistribution([
      [53, 0.3],
      [11, 0.4],
      [32, 0.3],
    ]);
    expect(r.distribution!.outcomes.map((o) => o.x)).toEqual([11, 32, 53]);
    expect(r.issues.some((i) => /sorted/.test(i.message))).toBe(true);
  });

  it("4. combines duplicate outcomes and warns", () => {
    const r = parseDistribution([
      [10, 0.1],
      [10, 0.2],
      [20, 0.7],
    ]);
    expect(r.ok).toBe(true);
    expect(r.distribution!.outcomes).toEqual([
      { x: 10, p: 0.3 },
      { x: 20, p: 0.7 },
    ]);
    expect(r.issues.some((i) => i.severity === "warning" && /appears 2 times/.test(i.message))).toBe(true);
  });

  it("5. keeps zero-probability outcomes with a warning", () => {
    const r = parseDistribution([
      [1, 0],
      [2, 1],
    ]);
    expect(r.ok).toBe(true);
    expect(r.distribution!.outcomes).toHaveLength(2);
    expect(r.issues.some((i) => i.severity === "warning" && /zero probability/.test(i.message))).toBe(true);
  });

  it("6. accepts negative outcomes", () => {
    const r = parseDistribution([
      [-100, 0.5],
      [-5.5, 0.5],
    ]);
    expect(r.ok).toBe(true);
    expect(r.distribution!.outcomes[0].x).toBe(-100);
  });

  it("7. reads percentages: formatted decimals, '40%' text, and percentage points", () => {
    // Excel cells formatted as % arrive as decimals already.
    expect(
      parseDistribution([
        [1, 0.4],
        [2, 0.6],
      ]).ok,
    ).toBe(true);
    const text = parseDistribution([
      [11, "40%"],
      [32, "30%"],
      [53, "30 %"],
    ]);
    expect(text.ok).toBe(true);
    expect(text.distribution!.outcomes.map((o) => o.p)).toEqual([0.4, 0.3, 0.3]);

    const points = parseDistribution([
      [11, 40],
      [32, 30],
      [53, 30],
    ]);
    expect(points.ok).toBe(true);
    expect(points.probabilityFormat).toBe("percent");
    expect(points.distribution!.outcomes.map((o) => o.p)).toEqual([0.4, 0.3, 0.3]);

    const explicit = parseDistribution([[5, 100]], { probabilityFormat: "percent" });
    expect(explicit.ok).toBe(true);
    expect(explicit.distribution!.outcomes[0].p).toBe(1);

    const decimalOnly = parseDistribution(
      [
        [11, 40],
        [32, 60],
      ],
      { probabilityFormat: "decimal" },
    );
    expect(decimalOnly.ok).toBe(false);
  });

  it("8. rejects probabilities that do not sum to 1", () => {
    const r = parseDistribution([
      [1, 0.5],
      [2, 0.4],
    ]);
    expect(r.ok).toBe(false);
    expect(r.distribution).toBeNull();
    expect(errors(r)[0].message).toMatch(/sum to 0\.9, not 1/);
  });

  it("8b. respects a configurable tolerance", () => {
    const values = [
      [1, 0.333],
      [2, 0.333],
      [3, 0.333],
    ];
    expect(parseDistribution(values).ok).toBe(false);
    expect(parseDistribution(values, { tolerance: 0.01 }).ok).toBe(true);
  });

  it("rejects probabilities outside [0, 1]", () => {
    const r = parseDistribution([
      [1, -0.2],
      [2, 1.2],
    ]);
    expect(errors(r).map((e) => e.cell)).toEqual(["B1", "B2"]);
  });

  it("9. single outcome with probability 1", () => {
    const r = parseDistribution([[42, 1]]);
    expect(r.ok).toBe(true);
    expect(r.distribution!.outcomes).toEqual([{ x: 42, p: 1 }]);
  });

  it("10. accepts sums that are 1 up to floating-point noise", () => {
    const r = parseDistribution(Array.from({ length: 10 }, (_, i) => [i, 0.1]));
    expect(r.ok).toBe(true);
    expect(Math.abs(r.probabilitySum - 1)).toBeLessThan(1e-12);
  });

  it("reports non-numeric cells with sheet addresses relative to the origin", () => {
    const r = parseDistribution(
      [
        ["Outcome", "P"],
        ["abc", 0.5],
        [2, "#DIV/0!"],
        [3, null],
      ],
      { origin: { rowIndex: 4, columnIndex: 2 } },
    );
    expect(errors(r).map((e) => e.cell)).toEqual(["C6", "D7", "D8"]);
  });

  it("ignores fully blank rows and requires exactly two columns", () => {
    const r = parseDistribution([
      [1, 0.5],
      ["", ""],
      [2, 0.5],
    ]);
    expect(r.ok).toBe(true);
    expect(r.issues.some((i) => /blank row/.test(i.message))).toBe(true);
    expect(parseDistribution([[1, 0.5, 3]]).ok).toBe(false);
    expect(parseDistribution([[1]]).ok).toBe(false);
    expect(parseDistribution([]).ok).toBe(false);
    expect(parseDistribution([["Outcome", "P"]]).ok).toBe(false);
  });
});
