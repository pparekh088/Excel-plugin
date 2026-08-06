/**
 * AI.FORECAST — Holt-Winters / exponential smoothing (handoff §8).
 *
 * Deliberately statistical, not generative: the numbers come from ETS, and an
 * LLM is only ever asked to narrate them. A forecast a model invented would be
 * indistinguishable from one it computed, and in a financial model that
 * distinction is the whole point.
 *
 * Implements simple, Holt's linear, and additive Holt-Winters, choosing by
 * what the data supports, and reports which method it used so the caller can
 * say so.
 */

export type ForecastMethod = "simple" | "holt" | "holt-winters";

export interface ForecastResult {
  /** Forecast values, one per horizon step. */
  values: number[];
  method: ForecastMethod;
  /** Chosen smoothing parameters. */
  alpha: number;
  beta?: number;
  gamma?: number;
  seasonLength?: number;
  /** In-sample mean absolute error — a rough confidence signal. */
  mae: number;
  /** Stated honestly when the series is too short for the method asked for. */
  caveat?: string;
}

function simpleExponential(series: number[], alpha: number): { fitted: number[]; level: number } {
  let level = series[0] ?? 0;
  const fitted: number[] = [level];
  for (let i = 1; i < series.length; i++) {
    level = alpha * series[i]! + (1 - alpha) * level;
    fitted.push(level);
  }
  return { fitted, level };
}

function holt(
  series: number[],
  alpha: number,
  beta: number
): { fitted: number[]; level: number; trend: number } {
  let level = series[0] ?? 0;
  let trend = (series[1] ?? level) - level;
  const fitted: number[] = [level];
  for (let i = 1; i < series.length; i++) {
    const previousLevel = level;
    level = alpha * series[i]! + (1 - alpha) * (level + trend);
    trend = beta * (level - previousLevel) + (1 - beta) * trend;
    fitted.push(level + trend);
  }
  return { fitted, level, trend };
}

function holtWinters(
  series: number[],
  alpha: number,
  beta: number,
  gamma: number,
  season: number
): { fitted: number[]; level: number; trend: number; seasonal: number[] } {
  const seasonal: number[] = [];
  const firstCycle = series.slice(0, season);
  const firstMean = firstCycle.reduce((sum, value) => sum + value, 0) / season;
  for (let i = 0; i < season; i++) seasonal.push((series[i] ?? firstMean) - firstMean);

  let level = firstMean;
  let trend =
    (series.slice(season, season * 2).reduce((sum, value) => sum + value, 0) / season - firstMean) /
    season;
  const fitted: number[] = [];

  for (let i = 0; i < series.length; i++) {
    const seasonIndex = i % season;
    const value = series[i]!;
    const previousLevel = level;
    level = alpha * (value - seasonal[seasonIndex]!) + (1 - alpha) * (level + trend);
    trend = beta * (level - previousLevel) + (1 - beta) * trend;
    seasonal[seasonIndex] = gamma * (value - level) + (1 - gamma) * seasonal[seasonIndex]!;
    fitted.push(level + trend + seasonal[seasonIndex]!);
  }
  return { fitted, level, trend, seasonal };
}

function meanAbsoluteError(actual: number[], fitted: number[]): number {
  let total = 0;
  let count = 0;
  for (let i = 0; i < actual.length; i++) {
    const predicted = fitted[i];
    if (predicted === undefined) continue;
    total += Math.abs(actual[i]! - predicted);
    count++;
  }
  return count === 0 ? 0 : total / count;
}

export interface ForecastOptions {
  horizon: number;
  /** Periods per season (12 monthly, 4 quarterly). 0 or 1 disables seasonality. */
  seasonLength?: number;
  method?: ForecastMethod | "auto";
}

/**
 * Grid search over smoothing parameters, minimising in-sample MAE. Small and
 * fixed — this runs server-side per call and must stay predictable.
 */
export function forecast(series: number[], options: ForecastOptions): ForecastResult {
  const clean = series.filter((value) => Number.isFinite(value));
  const horizon = Math.max(1, Math.floor(options.horizon));

  if (clean.length === 0) {
    return { values: new Array(horizon).fill(0), method: "simple", alpha: 0, mae: 0, caveat: "No numeric history was supplied; the forecast is all zeros." };
  }
  if (clean.length < 3) {
    const last = clean[clean.length - 1]!;
    return {
      values: new Array(horizon).fill(last),
      method: "simple",
      alpha: 1,
      mae: 0,
      caveat: `Only ${clean.length} data point(s): the forecast simply repeats the last value.`,
    };
  }

  const season = options.seasonLength ?? 0;
  const wantsSeasonal =
    (options.method === "holt-winters" || options.method === "auto" || options.method === undefined) &&
    season >= 2 &&
    clean.length >= season * 2;

  const grid = [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9];
  let best: ForecastResult | null = null;

  const consider = (candidate: ForecastResult): void => {
    if (!best || candidate.mae < best.mae) best = candidate;
  };

  if (wantsSeasonal) {
    for (const alpha of grid) {
      for (const beta of grid) {
        for (const gamma of [0.1, 0.3, 0.5, 0.7]) {
          const model = holtWinters(clean, alpha, beta, gamma, season);
          const values: number[] = [];
          for (let step = 1; step <= horizon; step++) {
            const seasonIndex = (clean.length + step - 1) % season;
            values.push(model.level + step * model.trend + model.seasonal[seasonIndex]!);
          }
          consider({
            values,
            method: "holt-winters",
            alpha,
            beta,
            gamma,
            seasonLength: season,
            mae: meanAbsoluteError(clean, model.fitted),
          });
        }
      }
    }
  }

  if (options.method === "holt" || options.method === "auto" || options.method === undefined || !best) {
    for (const alpha of grid) {
      for (const beta of grid) {
        const model = holt(clean, alpha, beta);
        const values: number[] = [];
        for (let step = 1; step <= horizon; step++) {
          values.push(model.level + step * model.trend);
        }
        consider({
          values,
          method: "holt",
          alpha,
          beta,
          mae: meanAbsoluteError(clean, model.fitted),
        });
      }
    }
  }

  if (options.method === "simple" || !best) {
    for (const alpha of grid) {
      const model = simpleExponential(clean, alpha);
      consider({
        values: new Array(horizon).fill(model.level),
        method: "simple",
        alpha,
        mae: meanAbsoluteError(clean, model.fitted),
      });
    }
  }

  const result = best as ForecastResult | null;
  if (!result) {
    const last = clean[clean.length - 1]!;
    return { values: new Array(horizon).fill(last), method: "simple", alpha: 1, mae: 0 };
  }

  if (options.seasonLength && options.seasonLength >= 2 && result.method !== "holt-winters") {
    result.caveat =
      `Seasonality of ${options.seasonLength} was requested but the series has only ` +
      `${clean.length} points — at least ${options.seasonLength * 2} are needed, so a ` +
      `non-seasonal model was used instead.`;
  }
  return result;
}
