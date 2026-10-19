import { describe, expect, it } from "vitest";
import { automaticStart, buildStepCoordinates, cumulativeProbabilities, resolveXRange } from "../src/core/cdf";
import { parseDistribution } from "../src/core/parse";
import { Distribution, Point } from "../src/core/types";

const MAKE = [
  [11, 0.4],
  [32, 0.3],
  [53, 0.3],
];
const BUY = [
  [15, 0.4],
  [30, 0.3],
  [45, 0.3],
];

function dist(values: (string | number)[][]): Distribution {
  const parsed = parseDistribution(values);
  expect(parsed.issues.filter((i) => i.severity === "error")).toEqual([]);
  return parsed.distribution!;
}

const pts = (pairs: number[][]): Point[] => pairs.map(([x, y]) => ({ x, y }));

describe("buildStepCoordinates", () => {
  it("1. Make Component example matches the specification exactly", () => {
    expect(buildStepCoordinates(dist(MAKE), { xStart: 0 })).toEqual(
      pts([
        [0, 0],
        [11, 0],
        [11, 0.4],
        [32, 0.4],
        [32, 0.7],
        [53, 0.7],
        [53, 1],
      ]),
    );
  });

  it("2. Buy Component example", () => {
    expect(buildStepCoordinates(dist(BUY), { xStart: 0 })).toEqual(
      pts([
        [0, 0],
        [15, 0],
        [15, 0.4],
        [30, 0.4],
        [30, 0.7],
        [45, 0.7],
        [45, 1],
      ]),
    );
  });

  it("3. unsorted outcomes are sorted first", () => {
    const unsorted = [
      [53, 0.3],
      [11, 0.4],
      [32, 0.3],
    ];
    expect(buildStepCoordinates(dist(unsorted), { xStart: 0 })).toEqual(
      buildStepCoordinates(dist(MAKE), { xStart: 0 }),
    );
  });

  it("4. duplicate outcomes are merged into one jump", () => {
    const dup = [
      [11, 0.2],
      [32, 0.3],
      [11, 0.2],
      [53, 0.3],
    ];
    expect(buildStepCoordinates(dist(dup), { xStart: 0 })).toEqual(buildStepCoordinates(dist(MAKE), { xStart: 0 }));
  });

  it("5. zero-probability outcomes add no step", () => {
    const withZero = [
      [11, 0.4],
      [20, 0],
      [32, 0.3],
      [53, 0.3],
      [60, 0],
    ];
    expect(buildStepCoordinates(dist(withZero), { xStart: 0 })).toEqual(
      buildStepCoordinates(dist(MAKE), { xStart: 0 }),
    );
  });

  it("6. negative outcomes work with a start at or below the minimum", () => {
    const neg = dist([
      [-20, 0.25],
      [0, 0.25],
      [50, 0.5],
    ]);
    const xStart = automaticStart(-20, 50);
    expect(xStart).toBeLessThan(-20);
    expect(buildStepCoordinates(neg, { xStart })).toEqual(
      pts([
        [xStart, 0],
        [-20, 0],
        [-20, 0.25],
        [0, 0.25],
        [0, 0.5],
        [50, 0.5],
        [50, 1],
      ]),
    );
    expect(() => buildStepCoordinates(neg, { xStart: 0 })).toThrow(/greater than the smallest outcome/);
  });

  it("drops the duplicate first point when xStart equals the first outcome", () => {
    expect(buildStepCoordinates(dist(MAKE), { xStart: 11 })).toEqual(
      pts([
        [11, 0],
        [11, 0.4],
        [32, 0.4],
        [32, 0.7],
        [53, 0.7],
        [53, 1],
      ]),
    );
  });

  it("extends the final segment to xEnd at probability 1", () => {
    const out = buildStepCoordinates(dist(MAKE), { xStart: 0, xEnd: 60 });
    expect(out[out.length - 1]).toEqual({ x: 60, y: 1 });
    expect(out).toHaveLength(8);
    // xEnd equal to the last outcome adds nothing
    expect(buildStepCoordinates(dist(MAKE), { xStart: 0, xEnd: 53 })).toHaveLength(7);
    expect(() => buildStepCoordinates(dist(MAKE), { xStart: 0, xEnd: 40 })).toThrow(/less than the largest/);
  });

  it("9. single-outcome distribution is a single vertical jump", () => {
    expect(buildStepCoordinates(dist([[7, 1]]), { xStart: 0 })).toEqual(
      pts([
        [0, 0],
        [7, 0],
        [7, 1],
      ]),
    );
  });

  it("10. floating-point noise is removed and the CDF ends at exactly 1", () => {
    const tenths = dist(Array.from({ length: 10 }, (_, i) => [i + 1, 0.1]));
    const out = buildStepCoordinates(tenths, { xStart: 0 });
    expect(out.map((p) => p.y)).toContain(0.3);
    expect(out.map((p) => p.y)).toContain(0.7);
    expect(out[out.length - 1].y).toBe(1);
    // Every y is non-decreasing.
    out.forEach((p, i) => i > 0 && expect(p.y).toBeGreaterThanOrEqual(out[i - 1].y));

    const thirds = dist([
      [1, 1 / 3],
      [2, 1 / 3],
      [3, 1 / 3],
    ]);
    const t = buildStepCoordinates(thirds, { xStart: 0 });
    expect(t[t.length - 1].y).toBe(1);
  });

  it("produces horizontal segments and vertical jumps only", () => {
    const out = buildStepCoordinates(dist(BUY), { xStart: 0, xEnd: 50 });
    for (let i = 1; i < out.length; i++) {
      const horizontal = out[i].y === out[i - 1].y;
      const vertical = out[i].x === out[i - 1].x;
      expect(horizontal || vertical).toBe(true);
    }
  });
});

describe("cumulativeProbabilities", () => {
  it("snaps a final value within tolerance to 1", () => {
    const c = cumulativeProbabilities({
      outcomes: [
        { x: 1, p: 0.3333333 },
        { x: 2, p: 0.3333333 },
        { x: 3, p: 0.3333334 },
      ],
    });
    expect(c[2].F).toBe(1);
  });
});

describe("resolveXRange", () => {
  const make = dist(MAKE);
  const buy = dist(BUY);
  const neg = dist([
    [-5, 0.5],
    [5, 0.5],
  ]);

  it("auto: zero for non-negative data, below the minimum otherwise", () => {
    expect(resolveXRange([make, buy], { startMode: "auto", endMode: "none" })).toEqual({ xStart: 0, xEnd: null });
    expect(resolveXRange([neg], { startMode: "auto", endMode: "none" }).xStart).toBeLessThan(-5);
  });

  it("chart minimum uses the smallest outcome over all distributions", () => {
    expect(resolveXRange([make, buy], { startMode: "min", endMode: "none" }).xStart).toBe(11);
  });

  it("chart maximum extends every CDF to the common largest outcome", () => {
    expect(resolveXRange([make, buy], { startMode: "zero", endMode: "max" })).toEqual({ xStart: 0, xEnd: 53 });
  });

  it("validates custom values", () => {
    expect(resolveXRange([make], { startMode: "custom", customStart: -10, endMode: "custom", customEnd: 100 })).toEqual({
      xStart: -10,
      xEnd: 100,
    });
    expect(() => resolveXRange([make], { startMode: "custom", customStart: 20, endMode: "none" })).toThrow();
    expect(() => resolveXRange([make], { startMode: "custom", endMode: "none" })).toThrow(/numeric/);
    expect(() => resolveXRange([make], { startMode: "zero", endMode: "custom", customEnd: 50 })).toThrow();
    expect(() => resolveXRange([neg], { startMode: "zero", endMode: "none" })).toThrow(/greater than the smallest/);
  });
});
