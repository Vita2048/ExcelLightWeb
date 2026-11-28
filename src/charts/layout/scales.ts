import { formatNumber } from '../../core/formats/cell-format';

export interface AxisScaleOptions {
  readonly includeZero?: boolean;
  readonly isPercent100?: boolean;
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  readonly targetTicks?: number;
}

export interface AxisScaleResult {
  readonly min: number;
  readonly max: number;
  readonly step: number;
  readonly ticks: readonly number[];
}

function niceNum(range: number, round: boolean): number {
  if (!Number.isFinite(range) || range <= 0) return 1;
  const exponent = Math.floor(Math.log10(range));
  const fraction = range / 10 ** exponent;
  let niceFraction: number;
  if (round) {
    if (fraction < 1.5) niceFraction = 1;
    else if (fraction < 3) niceFraction = 2;
    else if (fraction < 7) niceFraction = 5;
    else niceFraction = 10;
  } else {
    if (fraction <= 1) niceFraction = 1;
    else if (fraction <= 2) niceFraction = 2;
    else if (fraction <= 5) niceFraction = 5;
    else niceFraction = 10;
  }
  return niceFraction * 10 ** exponent;
}

export function computeNiceAxisScale(
  rawMin: number,
  rawMax: number,
  options: AxisScaleOptions = {},
): AxisScaleResult {
  if (options.isPercent100) {
    const min = options.min ?? (rawMin < 0 ? -1 : 0);
    const max = options.max ?? (rawMax > 0 ? 1 : 0);
    const step = options.majorUnit && options.majorUnit > 0 ? options.majorUnit : 0.2;
    const ticks: number[] = [];
    for (let v = min; v <= max + step * 0.001; v += step) {
      ticks.push(Number(v.toFixed(6)));
    }
    return { min, max, step, ticks };
  }

  let lo = Number.isFinite(rawMin) ? rawMin : 0;
  let hi = Number.isFinite(rawMax) ? rawMax : 0;

  if (options.includeZero) {
    if (lo > 0) lo = 0;
    if (hi < 0) hi = 0;
  }

  if (lo === hi) {
    if (lo === 0) {
      lo = 0;
      hi = 1;
    } else if (lo > 0) {
      lo = options.includeZero ? 0 : lo * 0.8;
      hi = hi * 1.2;
    } else {
      lo = lo * 1.2;
      hi = options.includeZero ? 0 : hi * 0.8;
    }
  }

  if (options.min !== undefined && Number.isFinite(options.min)) lo = options.min;
  if (options.max !== undefined && Number.isFinite(options.max)) hi = options.max;
  if (hi <= lo) hi = lo + 1;

  const targetTicks = Math.max(2, options.targetTicks ?? 6);
  const range = niceNum(hi - lo, false);
  const step = options.majorUnit && options.majorUnit > 0
    ? options.majorUnit
    : niceNum(range / (targetTicks - 1), true);

  const niceMin = options.min !== undefined
    ? options.min
    : Math.floor(lo / step) * step;
  const niceMax = options.max !== undefined
    ? options.max
    : Math.ceil(hi / step) * step;

  const ticks: number[] = [];
  const maxIterations = 64;
  for (let i = 0; i < maxIterations; i++) {
    const value = niceMin + i * step;
    if (value > niceMax + step * 0.001) break;
    ticks.push(Math.abs(value) < step * 1e-9 ? 0 : Number(value.toPrecision(12)));
  }

  return {
    min: niceMin,
    max: niceMax,
    step,
    ticks: ticks.length >= 2 ? ticks : [niceMin, niceMax],
  };
}

export function formatAxisTick(
  value: number,
  numberFormat?: string,
  isPercent100 = false,
): string {
  if (isPercent100) {
    return `${Math.round(value * 100)}%`;
  }
  if (numberFormat && !/^general$/iu.test(numberFormat)) {
    return formatNumber(value, numberFormat).text;
  }
  return formatNumber(value).text;
}
