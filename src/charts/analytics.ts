import type { ChartBinMode, ChartErrorBars, ChartQuartileMethod, ChartTrendline } from '../core/chart';
import { formatNumber } from '../core/formats/cell-format';

export interface HistogramBin {
  readonly label: string;
  readonly lower: number;
  readonly upper: number;
  readonly count: number;
  readonly isUnderflow?: boolean;
  readonly isOverflow?: boolean;
}

export interface ParetoBin extends HistogramBin {
  readonly cumulativePercent: number;
}

export interface HistogramOptions {
  readonly binMode?: ChartBinMode;
  readonly binWidth?: number;
  readonly binCount?: number;
  readonly overflowBin?: number;
  readonly underflowBin?: number;
}

export interface BoxWhiskerStats {
  readonly min: number;
  readonly max: number;
  readonly q1: number;
  readonly median: number;
  readonly q3: number;
  readonly iqr: number;
  readonly mean: number;
  readonly whiskerLow: number;
  readonly whiskerHigh: number;
  readonly outliers: readonly number[];
  readonly innerPoints: readonly number[];
}

export interface TrendlinePoint {
  readonly x: number;
  readonly y: number;
}

export interface TrendlineResult {
  readonly points: readonly TrendlinePoint[];
  readonly rSquared?: number;
  readonly equationText?: string;
}

export interface ErrorBarDelta {
  readonly plus: number;
  readonly minus: number;
}

function formatEdge(v: number): string {
  return formatNumber(Number(v.toPrecision(6))).text;
}

export function computeHistogramBins(
  rawValues: readonly number[],
  options: HistogramOptions = {},
): readonly HistogramBin[] {
  const values = rawValues.filter(v => Number.isFinite(v));
  if (values.length === 0) return [];

  const minVal = Math.min(...values);
  const maxVal = Math.max(...values);

  if (minVal === maxVal) {
    return [{
      label: `[${formatEdge(minVal)}, ${formatEdge(maxVal)}]`,
      lower: minVal,
      upper: maxVal,
      count: values.length,
    }];
  }

  const under = options.underflowBin !== undefined && Number.isFinite(options.underflowBin)
    ? options.underflowBin
    : undefined;
  const over = options.overflowBin !== undefined && Number.isFinite(options.overflowBin)
    ? options.overflowBin
    : undefined;

  const coreStart = under !== undefined ? under : minVal;
  const coreEnd = over !== undefined ? Math.max(coreStart + 1e-9, over) : maxVal;

  let binCount: number;
  let binWidth: number;

  if (options.binMode === 'binWidth' && options.binWidth && options.binWidth > 0) {
    binWidth = options.binWidth;
    binCount = Math.max(1, Math.min(200, Math.ceil((coreEnd - coreStart) / binWidth)));
  } else if (options.binMode === 'binCount' && options.binCount && options.binCount >= 1) {
    binCount = Math.max(1, Math.min(200, Math.floor(options.binCount)));
    binWidth = (coreEnd - coreStart) / binCount;
  } else {
    // Scott's rule matching Excel's automatic histogram bin width
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    const variance = values.reduce((a, b) => a + (b - mean) ** 2, 0) / Math.max(1, values.length - 1);
    const stdDev = Math.sqrt(variance);
    const scott = stdDev > 0 ? (3.5 * stdDev) / Math.cbrt(values.length) : 0;
    const rawSpan = coreEnd - coreStart;
    if (scott > 0 && rawSpan / scott >= 1 && rawSpan / scott <= 50) {
      binWidth = scott;
      binCount = Math.max(1, Math.min(50, Math.ceil(rawSpan / binWidth)));
      binWidth = rawSpan / binCount;
    } else {
      binCount = Math.max(1, Math.min(20, Math.ceil(Math.cbrt(values.length) * 2)));
      binWidth = rawSpan / binCount;
    }
  }

  const bins: { label: string; lower: number; upper: number; count: number; isUnderflow?: boolean; isOverflow?: boolean }[] = [];
  if (under !== undefined) {
    bins.push({
      label: `≤ ${formatEdge(under)}`,
      lower: -Infinity,
      upper: under,
      count: 0,
      isUnderflow: true,
    });
  }

  for (let i = 0; i < binCount; i++) {
    const lower = coreStart + i * binWidth;
    const upper = i === binCount - 1 && over === undefined ? coreEnd : coreStart + (i + 1) * binWidth;
    const openLeft = i > 0 || under !== undefined;
    bins.push({
      label: `${openLeft ? '(' : '['}${formatEdge(lower)}, ${formatEdge(upper)}]`,
      lower,
      upper,
      count: 0,
    });
  }

  if (over !== undefined) {
    bins.push({
      label: `> ${formatEdge(over)}`,
      lower: over,
      upper: Infinity,
      count: 0,
      isOverflow: true,
    });
  }

  for (const v of values) {
    if (under !== undefined && v <= under) {
      bins[0]!.count++;
      continue;
    }
    if (over !== undefined && v > over) {
      bins[bins.length - 1]!.count++;
      continue;
    }
    const startIdx = under !== undefined ? 1 : 0;
    const endIdx = over !== undefined ? bins.length - 2 : bins.length - 1;
    let placed = false;
    for (let i = startIdx; i <= endIdx; i++) {
      const b = bins[i]!;
      const isFirst = i === 0;
      const inBin = isFirst ? v >= b.lower && v <= b.upper : v > b.lower && v <= b.upper;
      if (inBin) {
        b.count++;
        placed = true;
        break;
      }
    }
    if (!placed && endIdx >= startIdx) {
      bins[v <= coreStart ? startIdx : endIdx]!.count++;
    }
  }

  return bins;
}

export function computeParetoBins(
  rawValues: readonly number[],
  options: HistogramOptions = {},
): readonly ParetoBin[] {
  const bins = computeHistogramBins(rawValues, options);
  const sorted = [...bins].sort((a, b) => b.count - a.count);
  const total = sorted.reduce((sum, b) => sum + b.count, 0);
  let running = 0;
  return sorted.map(b => {
    running += b.count;
    return {
      ...b,
      cumulativePercent: total > 0 ? running / total : 0,
    };
  });
}

function interpolateQuartile(sorted: readonly number[], p: number, method: ChartQuartileMethod): number {
  const n = sorted.length;
  if (n === 0) return 0;
  if (n === 1) return sorted[0]!;
  if (method === 'exclusive') {
    const rank = p * (n + 1);
    if (rank <= 1) return sorted[0]!;
    if (rank >= n) return sorted[n - 1]!;
    const lo = Math.floor(rank) - 1;
    const frac = rank - Math.floor(rank);
    return sorted[lo]! + frac * (sorted[lo + 1]! - sorted[lo]!);
  }
  const rank = 1 + p * (n - 1);
  const lo = Math.floor(rank) - 1;
  const hi = Math.min(n - 1, lo + 1);
  const frac = rank - Math.floor(rank);
  return sorted[lo]! + frac * (sorted[hi]! - sorted[lo]!);
}

export function computeBoxWhiskerStats(
  rawValues: readonly number[],
  method: ChartQuartileMethod = 'exclusive',
): BoxWhiskerStats | undefined {
  const sorted = rawValues.filter(v => Number.isFinite(v)).sort((a, b) => a - b);
  if (sorted.length === 0) return undefined;

  const q1 = interpolateQuartile(sorted, 0.25, method);
  const median = interpolateQuartile(sorted, 0.5, method);
  const q3 = interpolateQuartile(sorted, 0.75, method);
  const iqr = Math.max(0, q3 - q1);
  const mean = sorted.reduce((a, b) => a + b, 0) / sorted.length;

  const lowFence = q1 - 1.5 * iqr;
  const highFence = q3 + 1.5 * iqr;

  const innerPoints: number[] = [];
  const outliers: number[] = [];
  for (const v of sorted) {
    if (v < lowFence || v > highFence) outliers.push(v);
    else innerPoints.push(v);
  }

  const whiskerLow = innerPoints.length > 0 ? innerPoints[0]! : sorted[0]!;
  const whiskerHigh = innerPoints.length > 0 ? innerPoints[innerPoints.length - 1]! : sorted[sorted.length - 1]!;

  return {
    min: sorted[0]!,
    max: sorted[sorted.length - 1]!,
    q1,
    median,
    q3,
    iqr,
    mean,
    whiskerLow,
    whiskerHigh,
    outliers,
    innerPoints,
  };
}

function solveLinearSystem(matrix: number[][], rhs: number[]): number[] | undefined {
  const n = rhs.length;
  const a = matrix.map((row, i) => [...row, rhs[i]!]);
  for (let col = 0; col < n; col++) {
    let pivotRow = col;
    let best = Math.abs(a[col]![col]!);
    for (let r = col + 1; r < n; r++) {
      const val = Math.abs(a[r]![col]!);
      if (val > best) {
        best = val;
        pivotRow = r;
      }
    }
    if (best < 1e-12) return undefined;
    if (pivotRow !== col) {
      [a[col], a[pivotRow]] = [a[pivotRow]!, a[col]!];
    }
    const pivot = a[col]![col]!;
    for (let j = col; j <= n; j++) a[col]![j]! /= pivot;
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const factor = a[r]![col]!;
      for (let j = col; j <= n; j++) {
        a[r]![j]! -= factor * a[col]![j]!;
      }
    }
  }
  return a.map(row => row[n]!);
}

function rSquared(actual: readonly number[], predicted: readonly number[]): number | undefined {
  if (actual.length < 2) return undefined;
  const meanY = actual.reduce((a, b) => a + b, 0) / actual.length;
  let ssTot = 0;
  let ssRes = 0;
  for (let i = 0; i < actual.length; i++) {
    ssTot += (actual[i]! - meanY) ** 2;
    ssRes += (actual[i]! - predicted[i]!) ** 2;
  }
  if (ssTot === 0) return 1;
  return Math.max(0, Math.min(1, 1 - ssRes / ssTot));
}

function fmtCoeff(v: number): string {
  if (Math.abs(v) < 1e-9) return '0';
  return Number(v.toPrecision(4)).toString();
}

export function computeTrendline(
  rawPoints: readonly TrendlinePoint[],
  config: ChartTrendline,
): TrendlineResult {
  const pts = rawPoints
    .filter(p => Number.isFinite(p.x) && Number.isFinite(p.y))
    .sort((a, b) => a.x - b.x);
  if (pts.length < 2) return { points: [] };

  const minX = pts[0]!.x - (config.backward ?? 0);
  const maxX = pts[pts.length - 1]!.x + (config.forward ?? 0);
  const steps = 40;

  if (config.type === 'movingAverage') {
    const period = Math.max(2, Math.min(pts.length, config.period ?? 2));
    const maPoints: TrendlinePoint[] = [];
    for (let i = period - 1; i < pts.length; i++) {
      let sum = 0;
      for (let k = i - period + 1; k <= i; k++) sum += pts[k]!.y;
      maPoints.push({ x: pts[i]!.x, y: sum / period });
    }
    return { points: maPoints };
  }

  if (config.type === 'linear') {
    let slope: number;
    let intercept: number;
    if (config.intercept !== undefined && Number.isFinite(config.intercept)) {
      intercept = config.intercept;
      let num = 0;
      let den = 0;
      for (const p of pts) {
        num += p.x * (p.y - intercept);
        den += p.x * p.x;
      }
      slope = den !== 0 ? num / den : 0;
    } else {
      const n = pts.length;
      const sumX = pts.reduce((a, p) => a + p.x, 0);
      const sumY = pts.reduce((a, p) => a + p.y, 0);
      const meanX = sumX / n;
      const meanY = sumY / n;
      let num = 0;
      let den = 0;
      for (const p of pts) {
        num += (p.x - meanX) * (p.y - meanY);
        den += (p.x - meanX) ** 2;
      }
      slope = den !== 0 ? num / den : 0;
      intercept = meanY - slope * meanX;
    }
    const evalFn = (x: number) => slope * x + intercept;
    const r2 = rSquared(pts.map(p => p.y), pts.map(p => evalFn(p.x)));
    const sign = intercept >= 0 ? `+ ${fmtCoeff(intercept)}` : `- ${fmtCoeff(Math.abs(intercept))}`;
    const eq = `y = ${fmtCoeff(slope)}x ${sign}`;
    return {
      points: [{ x: minX, y: evalFn(minX) }, { x: maxX, y: evalFn(maxX) }],
      ...(r2 !== undefined && { rSquared: r2 }),
      equationText: eq,
    };
  }

  if (config.type === 'exponential') {
    const valid = pts.filter(p => p.y > 0);
    if (valid.length < 2) return { points: [] };
    const lnY = valid.map(p => Math.log(p.y));
    const n = valid.length;
    const meanX = valid.reduce((a, p) => a + p.x, 0) / n;
    const meanLnY = lnY.reduce((a, b) => a + b, 0) / n;
    let num = 0;
    let den = 0;
    for (let i = 0; i < n; i++) {
      num += (valid[i]!.x - meanX) * (lnY[i]! - meanLnY);
      den += (valid[i]!.x - meanX) ** 2;
    }
    const b = den !== 0 ? num / den : 0;
    const a = config.intercept !== undefined && config.intercept > 0
      ? config.intercept
      : Math.exp(meanLnY - b * meanX);
    const evalFn = (x: number) => a * Math.exp(b * x);
    const out: TrendlinePoint[] = [];
    for (let i = 0; i <= steps; i++) {
      const x = minX + ((maxX - minX) * i) / steps;
      out.push({ x, y: evalFn(x) });
    }
    const r2 = rSquared(valid.map(p => p.y), valid.map(p => evalFn(p.x)));
    return {
      points: out,
      ...(r2 !== undefined && { rSquared: r2 }),
      equationText: `y = ${fmtCoeff(a)}e^(${fmtCoeff(b)}x)`,
    };
  }

  if (config.type === 'logarithmic') {
    const valid = pts.filter(p => p.x > 0);
    if (valid.length < 2) return { points: [] };
    const lnX = valid.map(p => Math.log(p.x));
    const n = valid.length;
    const meanLnX = lnX.reduce((a, b) => a + b, 0) / n;
    const meanY = valid.reduce((a, p) => a + p.y, 0) / n;
    let num = 0;
    let den = 0;
    for (let i = 0; i < n; i++) {
      num += (lnX[i]! - meanLnX) * (valid[i]!.y - meanY);
      den += (lnX[i]! - meanLnX) ** 2;
    }
    const a = den !== 0 ? num / den : 0;
    const b = meanY - a * meanLnX;
    const evalFn = (x: number) => a * Math.log(x) + b;
    const startX = Math.max(1e-4, minX);
    const out: TrendlinePoint[] = [];
    for (let i = 0; i <= steps; i++) {
      const x = startX + ((Math.max(startX, maxX) - startX) * i) / steps;
      out.push({ x, y: evalFn(x) });
    }
    const r2 = rSquared(valid.map(p => p.y), valid.map(p => evalFn(p.x)));
    const sign = b >= 0 ? `+ ${fmtCoeff(b)}` : `- ${fmtCoeff(Math.abs(b))}`;
    return {
      points: out,
      ...(r2 !== undefined && { rSquared: r2 }),
      equationText: `y = ${fmtCoeff(a)}ln(x) ${sign}`,
    };
  }

  if (config.type === 'power') {
    const valid = pts.filter(p => p.x > 0 && p.y > 0);
    if (valid.length < 2) return { points: [] };
    const lnX = valid.map(p => Math.log(p.x));
    const lnY = valid.map(p => Math.log(p.y));
    const n = valid.length;
    const meanLnX = lnX.reduce((a, b) => a + b, 0) / n;
    const meanLnY = lnY.reduce((a, b) => a + b, 0) / n;
    let num = 0;
    let den = 0;
    for (let i = 0; i < n; i++) {
      num += (lnX[i]! - meanLnX) * (lnY[i]! - meanLnY);
      den += (lnX[i]! - meanLnX) ** 2;
    }
    const b = den !== 0 ? num / den : 0;
    const a = Math.exp(meanLnY - b * meanLnX);
    const evalFn = (x: number) => a * (x ** b);
    const startX = Math.max(1e-4, minX);
    const out: TrendlinePoint[] = [];
    for (let i = 0; i <= steps; i++) {
      const x = startX + ((Math.max(startX, maxX) - startX) * i) / steps;
      out.push({ x, y: evalFn(x) });
    }
    const r2 = rSquared(valid.map(p => p.y), valid.map(p => evalFn(p.x)));
    return {
      points: out,
      ...(r2 !== undefined && { rSquared: r2 }),
      equationText: `y = ${fmtCoeff(a)}x^(${fmtCoeff(b)})`,
    };
  }

  // Polynomial order 2..6
  const order = Math.max(2, Math.min(6, config.polynomialOrder ?? 2));
  const degree = Math.min(order, pts.length - 1);
  const size = degree + 1;
  const mat: number[][] = Array.from({ length: size }, () => new Array<number>(size).fill(0));
  const rhs = new Array<number>(size).fill(0);
  for (let r = 0; r < size; r++) {
    for (let c = 0; c < size; c++) {
      let sum = 0;
      for (const p of pts) sum += p.x ** (r + c);
      mat[r]![c] = sum;
    }
    let ySum = 0;
    for (const p of pts) ySum += p.y * (p.x ** r);
    rhs[r] = ySum;
  }
  const coeffs = solveLinearSystem(mat, rhs);
  if (!coeffs) return { points: [] };
  const evalFn = (x: number) => coeffs.reduce((acc, c, pwr) => acc + c * (x ** pwr), 0);
  const out: TrendlinePoint[] = [];
  for (let i = 0; i <= steps; i++) {
    const x = minX + ((maxX - minX) * i) / steps;
    out.push({ x, y: evalFn(x) });
  }
  const r2 = rSquared(pts.map(p => p.y), pts.map(p => evalFn(p.x)));
  const terms = coeffs
    .map((c, pwr) => ({ c, pwr }))
    .reverse()
    .map(({ c, pwr }) => (pwr === 0 ? fmtCoeff(c) : pwr === 1 ? `${fmtCoeff(c)}x` : `${fmtCoeff(c)}x^${pwr}`))
    .join(' + ')
    .replaceAll('+ -', '- ');
  return {
    points: out,
    ...(r2 !== undefined && { rSquared: r2 }),
    equationText: `y = ${terms}`,
  };
}

export function computeErrorBars(
  values: readonly (number | null)[],
  config: ChartErrorBars,
  customPlus?: readonly (number | null)[],
  customMinus?: readonly (number | null)[],
): readonly ErrorBarDelta[] {
  const finite = values.filter((v): v is number => v !== null && Number.isFinite(v));
  const mean = finite.length > 0 ? finite.reduce((a, b) => a + b, 0) / finite.length : 0;
  const sampleVar = finite.length > 1
    ? finite.reduce((a, b) => a + (b - mean) ** 2, 0) / (finite.length - 1)
    : 0;
  const stdDev = Math.sqrt(sampleVar);
  const stdErr = finite.length > 0 ? stdDev / Math.sqrt(finite.length) : 0;

  return values.map((val, idx) => {
    if (val === null || !Number.isFinite(val)) return { plus: 0, minus: 0 };
    let magPlus = 0;
    let magMinus = 0;
    switch (config.type) {
      case 'fixed': {
        const v = Math.max(0, config.value ?? 1);
        magPlus = v;
        magMinus = v;
        break;
      }
      case 'percentage': {
        const pct = Math.max(0, config.value ?? 5) / 100;
        const delta = Math.abs(val) * pct;
        magPlus = delta;
        magMinus = delta;
        break;
      }
      case 'stdDev': {
        const mult = Math.max(0, config.value ?? 1);
        magPlus = stdDev * mult;
        magMinus = stdDev * mult;
        break;
      }
      case 'stdErr': {
        magPlus = stdErr;
        magMinus = stdErr;
        break;
      }
      case 'custom': {
        magPlus = Math.max(0, customPlus?.[idx] ?? config.value ?? 0);
        magMinus = Math.max(0, customMinus?.[idx] ?? config.value ?? 0);
        break;
      }
    }
    return {
      plus: config.direction === 'minus' ? 0 : magPlus,
      minus: config.direction === 'plus' ? 0 : magMinus,
    };
  });
}
