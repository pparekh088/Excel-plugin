import { describe, expect, it } from "vitest";
import { forecast } from "../../src/aifn/forecast";

describe("AI.FORECAST (statistical, not generative)", () => {
  it("extrapolates a linear trend", () => {
    const series = [10, 20, 30, 40, 50, 60];
    const result = forecast(series, { horizon: 3 });
    expect(result.values).toHaveLength(3);
    // Next values should continue ~70, 80, 90.
    expect(result.values[0]).toBeGreaterThan(60);
    expect(result.values[2]).toBeGreaterThan(result.values[0]!);
    expect(result.method).toBe("holt");
  });

  it("holds a flat series flat", () => {
    const result = forecast([100, 100, 100, 100, 100, 100], { horizon: 2 });
    for (const value of result.values) expect(Math.abs(value - 100)).toBeLessThan(1);
  });

  it("captures seasonality when there are enough cycles", () => {
    // Two years of quarterly data with a clear seasonal pattern.
    const series = [100, 140, 90, 160, 110, 150, 100, 170];
    const result = forecast(series, { horizon: 4, seasonLength: 4 });
    expect(result.method).toBe("holt-winters");
    expect(result.seasonLength).toBe(4);
    // Q2 and Q4 are the peaks; the forecast should preserve that shape.
    expect(result.values[1]!).toBeGreaterThan(result.values[0]!);
    expect(result.values[3]!).toBeGreaterThan(result.values[2]!);
  });

  it("says so when the series is too short for the seasonality requested", () => {
    const result = forecast([100, 120, 110, 130, 115], { horizon: 2, seasonLength: 4 });
    expect(result.method).not.toBe("holt-winters");
    expect(result.caveat).toContain("at least 8");
  });

  it("handles a single data point honestly", () => {
    const result = forecast([42], { horizon: 3 });
    expect(result.values).toEqual([42, 42, 42]);
    expect(result.caveat).toContain("repeats the last value");
  });

  it("handles an empty series without throwing", () => {
    const result = forecast([], { horizon: 2 });
    expect(result.values).toEqual([0, 0]);
    expect(result.caveat).toContain("No numeric history");
  });

  it("ignores non-numeric entries", () => {
    const series = [10, Number.NaN, 20, Number.POSITIVE_INFINITY, 30, 40];
    const result = forecast(series, { horizon: 1 });
    expect(Number.isFinite(result.values[0])).toBe(true);
  });

  it("reports the parameters it chose so a reviewer can check them", () => {
    const result = forecast([1, 2, 3, 4, 5, 6, 7, 8], { horizon: 2 });
    expect(result.alpha).toBeGreaterThan(0);
    expect(result.mae).toBeGreaterThanOrEqual(0);
    expect(result.method).toBeTruthy();
  });

  it("fits noisy data with a lower error than a naive flat line", () => {
    const series = [10, 13, 11, 16, 15, 19, 18, 23, 21, 26];
    const result = forecast(series, { horizon: 1 });
    const mean = series.reduce((sum, value) => sum + value, 0) / series.length;
    const flatMae =
      series.reduce((sum, value) => sum + Math.abs(value - mean), 0) / series.length;
    expect(result.mae).toBeLessThan(flatMae);
  });
});
