import { describe, expect, it } from "vitest";
import { compareDistributions, explainDominance, verdictFromDifferences } from "../src/core/dominance";
import { computeStats } from "../src/core/stats";
import { Distribution } from "../src/core/types";

const d = (pairs: [number, number][]): Distribution => ({
  outcomes: pairs.map(([x, p]) => ({ x, p })).sort((a, b) => a.x - b.x),
});

const MAKE = d([
  [11, 0.4],
  [32, 0.3],
  [53, 0.3],
]);
const BUY = d([
  [15, 0.4],
  [30, 0.3],
  [45, 0.3],
]);

describe("computeStats", () => {
  it("Make Component: EV, variance, SD", () => {
    const s = computeStats(MAKE);
    expect(s.expectedValue).toBe(29.9);
    // Var = 0.4*18.9^2 + 0.3*2.1^2 + 0.3*23.1^2 = 142.884 + 1.323 + 160.083
    expect(s.variance).toBeCloseTo(304.29, 10);
    expect(s.standardDeviation).toBeCloseTo(Math.sqrt(304.29), 10);
    expect(s.median).toBe(32);
    expect([s.min, s.max]).toEqual([11, 53]);
  });

  it("Buy Component: EV, variance, SD", () => {
    const s = computeStats(BUY);
    expect(s.expectedValue).toBe(28.5);
    // 0.4*13.5^2 + 0.3*1.5^2 + 0.3*16.5^2 = 72.9 + 0.675 + 81.675
    expect(s.variance).toBeCloseTo(155.25, 10);
  });

  it("single outcome has zero variance; zero-probability outcomes are ignored", () => {
    const s = computeStats(
      d([
        [5, 1],
        [9, 0],
      ]),
    );
    expect(s).toMatchObject({ expectedValue: 5, variance: 0, standardDeviation: 0, min: 5, max: 5, outcomes: 1 });
  });

  it("negative outcomes", () => {
    expect(
      computeStats(
        d([
          [-10, 0.5],
          [10, 0.5],
        ]),
      ),
    ).toMatchObject({ expectedValue: 0, variance: 100, standardDeviation: 10 });
  });
});

describe("verdictFromDifferences", () => {
  it("classifies sign patterns", () => {
    expect(verdictFromDifferences([0, -0.1, 0], 1e-9)).toBe("A");
    expect(verdictFromDifferences([0, 0.1], 1e-9)).toBe("B");
    expect(verdictFromDifferences([0.1, -0.1], 1e-9)).toBe("neither");
    expect(verdictFromDifferences([0, 1e-12], 1e-9)).toBe("identical");
  });
});

describe("compareDistributions: first-order dominance", () => {
  it("A FSD B when A's CDF is never above B's", () => {
    const A = d([
      [10, 0.5],
      [20, 0.5],
    ]);
    const B = d([
      [5, 0.5],
      [20, 0.5],
    ]);
    const r = compareDistributions(A, B);
    expect(r.fsd).toBe("A");
    expect(r.ssd).toBe("A"); // FSD implies SSD
    expect(compareDistributions(B, A).fsd).toBe("B");
  });

  it("evaluates at the union of outcomes including jumps", () => {
    const r = compareDistributions(MAKE, BUY);
    expect(r.rows.map((row) => row.x)).toEqual([11, 15, 30, 32, 45, 53]);
    expect(r.rows.map((row) => row.FA)).toEqual([0.4, 0.4, 0.4, 0.7, 0.7, 1]);
    expect(r.rows.map((row) => row.FB)).toEqual([0, 0.4, 0.7, 0.7, 1, 1]);
  });

  it("11a. crossing CDFs (Make vs Buy): neither FSD, and for 'higher is better' neither SSD", () => {
    const r = compareDistributions(MAKE, BUY);
    expect(r.fsd).toBe("neither");
    expect(r.cdfCrossings).toEqual([30]);
    // Integrated CDFs: G_Make - G_Buy = 0, 1.6, 1.6, 1.0, 1.0, -1.4 -> crosses.
    expect(r.rows.map((row) => +(row.IA - row.IB).toFixed(10))).toEqual([0, 1.6, 1.6, 1, 1, -1.4]);
    expect(r.ssd).toBe("neither");
    // Higher EV for Make does not make it dominant.
    expect(computeStats(MAKE).expectedValue).toBeGreaterThan(computeStats(BUY).expectedValue);
  });

  it("11b. as costs (lower is better), Buy second-order dominates Make but not first-order", () => {
    const r = compareDistributions(MAKE, BUY, { preference: "lower" });
    expect(r.fsd).toBe("neither");
    expect(r.ssd).toBe("B");
    // I(x) = ∫_x^∞ (1 - F): Make 18.9, 16.5, 7.5, 6.3, 2.4, 0; Buy 17.5, 13.5, 4.5, 3.9, 0, 0.
    expect(r.rows.map((row) => row.IA)).toEqual([18.9, 16.5, 7.5, 6.3, 2.4, 0]);
    expect(r.rows.map((row) => row.IB)).toEqual([17.5, 13.5, 4.5, 3.9, 0, 0]);
  });

  it("identical distributions (different input order)", () => {
    const r = compareDistributions(
      MAKE,
      d([
        [53, 0.3],
        [11, 0.4],
        [32, 0.3],
      ]),
    );
    expect(r.fsd).toBe("identical");
    expect(r.ssd).toBe("identical");
  });

  it("lower preference reverses a first-order ranking", () => {
    const A = d([
      [10, 0.5],
      [20, 0.5],
    ]);
    const B = d([
      [5, 0.5],
      [20, 0.5],
    ]);
    expect(compareDistributions(A, B, { preference: "lower" }).fsd).toBe("B");
  });
});

describe("compareDistributions: second-order dominance", () => {
  it("11c. a sure amount SSD a mean-preserving spread, without FSD", () => {
    const sure = d([[10, 1]]);
    const risky = d([
      [0, 0.5],
      [20, 0.5],
    ]);
    const r = compareDistributions(sure, risky);
    expect(r.fsd).toBe("neither");
    expect(r.ssd).toBe("A");
    expect(compareDistributions(risky, sure).ssd).toBe("B");
  });

  it("11d. higher mean and lower variance do not imply SSD", () => {
    const A = d([
      [-10, 0.01],
      [50, 0.99],
    ]);
    const B = d([
      [0, 0.5],
      [90, 0.5],
    ]);
    const sa = computeStats(A);
    const sb = computeStats(B);
    expect(sa.expectedValue).toBeGreaterThan(sb.expectedValue);
    expect(sa.variance).toBeLessThan(sb.variance);
    const r = compareDistributions(A, B);
    expect(r.fsd).toBe("neither");
    expect(r.ssd).toBe("neither");
    const text = explainDominance(r, "A", "B", { a: sa.expectedValue, b: sb.expectedValue });
    expect(text.ssdHeadline).toMatch(/Neither/);
    expect(text.notes.join(" ")).toMatch(/does not make it preferred/);
  });

  it("is exact at breakpoints with tiny differences", () => {
    const A = d([
      [0, 0.5],
      [1, 0.5],
    ]);
    const B = d([
      [0, 0.5],
      [1.000001, 0.5],
    ]);
    const r = compareDistributions(A, B, { tolerance: 1e-9 });
    expect(r.fsd).toBe("B");
  });
});

describe("explainDominance", () => {
  it("names the dominant distribution", () => {
    const r = compareDistributions(
      d([
        [10, 0.5],
        [20, 0.5],
      ]),
      d([
        [5, 0.5],
        [20, 0.5],
      ]),
    );
    const t = explainDominance(r, "Project X", "Project Y");
    expect(t.fsdHeadline).toBe("Project X first-order stochastically dominates Project Y.");
    expect(t.ssdDetail).toMatch(/Implied by first-order dominance/);
  });
});
