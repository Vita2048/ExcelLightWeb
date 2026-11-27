import type { FormulaExpression } from '../../core/formula-references';
import { error, finite, isArray, isError, isOmitted, isRange, type FormulaError, type Scalar } from '../formula-values';
import {
  aggregate, aggregateInputs, averageIf, collectNumbers, conditional, medianOf, modeOf,
  multiConditional, nthExtreme, percentileOf,
} from './helpers';
import type { FunctionContext, FunctionSignature, FunctionSpec } from './types';

const numbers: FunctionSignature = [['number1', '[number2]'], true];
const values: FunctionSignature = [['value1', '[value2]'], true];

function aggregateSpec(name: string, code: number, signature: FunctionSignature): FunctionSpec {
  return {
    name,
    category: 'statistical',
    minArgs: 1,
    maxArgs: 255,
    signature,
    evaluate(args, ctx) {
      if (!args.length) return error('#VALUE!');
      return aggregate(code, aggregateInputs(args, ctx, code));
    },
  };
}

function collectPairs(
  yExpr: FormulaExpression,
  xExpr: FormulaExpression,
  ctx: FunctionContext,
): { ys: number[]; xs: number[] } | FormulaError {
  const yVal = ctx.evaluate(yExpr, true);
  const xVal = ctx.evaluate(xExpr, true);
  if (isError(yVal)) return yVal;
  if (isError(xVal)) return xVal;
  const yArr = ctx.arrayOf(yVal);
  const xArr = ctx.arrayOf(xVal);
  if (!yArr || !xArr) return error('#VALUE!');
  if (yArr.rows * yArr.columns !== xArr.rows * xArr.columns) return error('#N/A');

  const ys: number[] = [];
  const xs: number[] = [];
  const total = yArr.rows * yArr.columns;
  for (let i = 0; i < total; i++) {
    const yr = Math.floor(i / yArr.columns), yc = i % yArr.columns;
    const xr = Math.floor(i / xArr.columns), xc = i % xArr.columns;
    const y = yArr.at(yr, yc);
    const x = xArr.at(xr, xc);
    if (isError(y)) return y;
    if (isError(x)) return x;
    if (typeof y === 'number' && typeof x === 'number') {
      ys.push(y);
      xs.push(x);
    }
  }
  return { ys, xs };
}

function linearRegression(ys: readonly number[], xs: readonly number[]): {
  slope: number;
  intercept: number;
  sumX2: number;
  sumY2: number;
  sumXY: number;
  n: number;
} | FormulaError {
  const n = ys.length;
  if (n < 1) return error('#DIV/0!');
  let meanX = 0, meanY = 0;
  for (let i = 0; i < n; i++) {
    meanX += xs[i]!;
    meanY += ys[i]!;
  }
  meanX /= n;
  meanY /= n;
  let sumX2 = 0, sumY2 = 0, sumXY = 0;
  for (let i = 0; i < n; i++) {
    const dx = xs[i]! - meanX;
    const dy = ys[i]! - meanY;
    sumX2 += dx * dx;
    sumY2 += dy * dy;
    sumXY += dx * dy;
  }
  if (sumX2 === 0) return error('#DIV/0!');
  const slope = sumXY / sumX2;
  const intercept = meanY - slope * meanX;
  return { slope, intercept, sumX2, sumY2, sumXY, n };
}

function rankImpl(avgTies: boolean, args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
  if (args.length < 2 || args.length > 3) return error('#VALUE!');
  const numVal = ctx.scalar(ctx.evaluate(args[0]!));
  if (isError(numVal)) return numVal;
  const target = ctx.number(numVal);
  if (!Number.isFinite(target)) return error('#VALUE!');

  let asc = false;
  if (args[2] && !isOmitted(ctx.evaluate(args[2]))) {
    const ordScalar = ctx.scalar(ctx.evaluate(args[2]));
    if (isError(ordScalar)) return ordScalar;
    const ordNum = ctx.number(ordScalar);
    if (!Number.isFinite(ordNum)) return error('#VALUE!');
    asc = ordNum !== 0;
  }

  const refVal = ctx.evaluate(args[1]!, true);
  if (isError(refVal)) return refVal;
  const nums: number[] = [];
  const push = (item: Scalar): FormulaError | undefined => {
    if (isError(item)) return item;
    if (typeof item === 'number') nums.push(item);
    return undefined;
  };
  if (isRange(refVal)) {
    for (const r of refVal) for (const v of ctx.values(r)) {
      const err = push(v);
      if (err) return err;
    }
  } else if (isArray(refVal)) {
    for (let r = 0; r < refVal.rows; r++) for (let c = 0; c < refVal.columns; c++) {
      const err = push(refVal.at(r, c));
      if (err) return err;
    }
  } else {
    return error('#VALUE!');
  }

  let before = 0;
  let equal = 0;
  for (const n of nums) {
    if (n === target) equal++;
    else if (asc ? n < target : n > target) before++;
  }
  if (equal === 0) return error('#N/A');
  return avgTies ? before + (equal + 1) / 2 : before + 1;
}

export const statisticalFunctions: readonly FunctionSpec[] = [
  aggregateSpec('AVERAGE', 1, numbers),
  aggregateSpec('MIN', 5, numbers),
  aggregateSpec('MAX', 4, numbers),
  aggregateSpec('COUNT', 2, values),
  aggregateSpec('COUNTA', 3, values),
  {
    name: 'COUNTIF',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['range', 'criteria']],
    evaluate: (args, ctx) => conditional('COUNTIF', args, ctx),
  },
  {
    name: 'COUNTIFS',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 254,
    signature: [['criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]'], true],
    evaluate: (args, ctx) => multiConditional('COUNTIFS', args, ctx),
  },
  {
    name: 'AVERAGEIF',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 3,
    signature: [['range', 'criteria', '[average_range]']],
    evaluate: averageIf,
  },
  {
    name: 'AVERAGEIFS',
    category: 'statistical',
    minArgs: 3,
    maxArgs: 255,
    signature: [['average_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]'], true],
    evaluate: (args, ctx) => multiConditional('AVERAGEIFS', args, ctx),
  },
  {
    name: 'MAXIFS',
    category: 'statistical',
    minArgs: 3,
    maxArgs: 255,
    signature: [['max_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]'], true],
    evaluate: (args, ctx) => multiConditional('MAXIFS', args, ctx),
  },
  {
    name: 'MINIFS',
    category: 'statistical',
    minArgs: 3,
    maxArgs: 255,
    signature: [['min_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]'], true],
    evaluate: (args, ctx) => multiConditional('MINIFS', args, ctx),
  },
  {
    name: 'COUNTBLANK',
    category: 'statistical',
    minArgs: 1,
    maxArgs: 1,
    signature: [['range']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 1) return error('#VALUE!');
      const val = ctx.evaluate(args[0]!, true);
      if (isError(val)) return val;
      if (!isRange(val) || val.length !== 1) return error('#VALUE!');
      const area = val[0]!;
      const total = (area.bottom - area.top + 1) * (area.right - area.left + 1);
      let nonEmpty = 0;
      for (const entry of ctx.entries(area)) {
        if (entry.value !== null && entry.value !== '') nonEmpty++;
      }
      return total - nonEmpty;
    },
  },
  {
    name: 'MEDIAN',
    category: 'statistical',
    minArgs: 1,
    maxArgs: 255,
    signature: numbers,
    evaluate(args, ctx): Scalar {
      const nums = collectNumbers(args, ctx);
      return isError(nums) ? nums : medianOf(nums);
    },
  },
  {
    name: 'MODE',
    category: 'statistical',
    minArgs: 1,
    maxArgs: 255,
    signature: numbers,
    evaluate(args, ctx): Scalar {
      const nums = collectNumbers(args, ctx);
      return isError(nums) ? nums : modeOf(nums);
    },
  },
  {
    name: 'MODE.SNGL',
    category: 'statistical',
    minArgs: 1,
    maxArgs: 255,
    signature: numbers,
    evaluate(args, ctx): Scalar {
      const nums = collectNumbers(args, ctx);
      return isError(nums) ? nums : modeOf(nums);
    },
  },
  aggregateSpec('STDEV.S', 7, numbers),
  aggregateSpec('STDEV.P', 8, numbers),
  aggregateSpec('VAR.S', 10, numbers),
  aggregateSpec('VAR.P', 11, numbers),
  {
    name: 'LARGE',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array', 'k']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 2) return error('#VALUE!');
      const nums = collectNumbers([args[0]!], ctx);
      if (isError(nums)) return nums;
      const kVal = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(kVal)) return kVal;
      const k = ctx.number(kVal);
      return Number.isNaN(k) ? error('#VALUE!') : nthExtreme(nums, k, true);
    },
  },
  {
    name: 'SMALL',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array', 'k']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 2) return error('#VALUE!');
      const nums = collectNumbers([args[0]!], ctx);
      if (isError(nums)) return nums;
      const kVal = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(kVal)) return kVal;
      const k = ctx.number(kVal);
      return Number.isNaN(k) ? error('#VALUE!') : nthExtreme(nums, k, false);
    },
  },
  {
    name: 'RANK',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 3,
    signature: [['number', 'ref', '[order]']],
    evaluate: (args, ctx) => rankImpl(false, args, ctx),
  },
  {
    name: 'RANK.EQ',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 3,
    signature: [['number', 'ref', '[order]']],
    evaluate: (args, ctx) => rankImpl(false, args, ctx),
  },
  {
    name: 'RANK.AVG',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 3,
    signature: [['number', 'ref', '[order]']],
    evaluate: (args, ctx) => rankImpl(true, args, ctx),
  },
  {
    name: 'PERCENTILE.INC',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array', 'k']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 2) return error('#VALUE!');
      const nums = collectNumbers([args[0]!], ctx);
      if (isError(nums)) return nums;
      const kVal = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(kVal)) return kVal;
      const k = ctx.number(kVal);
      return Number.isNaN(k) ? error('#VALUE!') : percentileOf(nums, k, true);
    },
  },
  {
    name: 'PERCENTILE.EXC',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array', 'k']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 2) return error('#VALUE!');
      const nums = collectNumbers([args[0]!], ctx);
      if (isError(nums)) return nums;
      const kVal = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(kVal)) return kVal;
      const k = ctx.number(kVal);
      return Number.isNaN(k) ? error('#VALUE!') : percentileOf(nums, k, false);
    },
  },
  {
    name: 'QUARTILE.INC',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array', 'quart']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 2) return error('#VALUE!');
      const nums = collectNumbers([args[0]!], ctx);
      if (isError(nums)) return nums;
      const qVal = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(qVal)) return qVal;
      const q = ctx.number(qVal);
      if (Number.isNaN(q)) return error('#VALUE!');
      const quart = Math.trunc(q);
      return quart < 0 || quart > 4 ? error('#NUM!') : percentileOf(nums, quart / 4, true);
    },
  },
  {
    name: 'QUARTILE.EXC',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array', 'quart']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 2) return error('#VALUE!');
      const nums = collectNumbers([args[0]!], ctx);
      if (isError(nums)) return nums;
      const qVal = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(qVal)) return qVal;
      const q = ctx.number(qVal);
      if (Number.isNaN(q)) return error('#VALUE!');
      const quart = Math.trunc(q);
      return quart < 1 || quart > 3 ? error('#NUM!') : percentileOf(nums, quart / 4, false);
    },
  },
  {
    name: 'CORREL',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array1', 'array2']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 2) return error('#VALUE!');
      const pair = collectPairs(args[0]!, args[1]!, ctx);
      if (isError(pair)) return pair;
      const reg = linearRegression(pair.ys, pair.xs);
      if (isError(reg)) return reg;
      if (reg.sumY2 === 0) return error('#DIV/0!');
      return finite(reg.sumXY / Math.sqrt(reg.sumX2 * reg.sumY2));
    },
  },
  {
    name: 'COVARIANCE.S',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array1', 'array2']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 2) return error('#VALUE!');
      const pair = collectPairs(args[0]!, args[1]!, ctx);
      if (isError(pair)) return pair;
      if (pair.ys.length < 2) return error('#DIV/0!');
      const reg = linearRegression(pair.ys, pair.xs);
      if (isError(reg)) return reg;
      return finite(reg.sumXY / (reg.n - 1));
    },
  },
  {
    name: 'SLOPE',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['known_ys', 'known_xs']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 2) return error('#VALUE!');
      const pair = collectPairs(args[0]!, args[1]!, ctx);
      if (isError(pair)) return pair;
      const reg = linearRegression(pair.ys, pair.xs);
      return isError(reg) ? reg : finite(reg.slope);
    },
  },
  {
    name: 'INTERCEPT',
    category: 'statistical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['known_ys', 'known_xs']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 2) return error('#VALUE!');
      const pair = collectPairs(args[0]!, args[1]!, ctx);
      if (isError(pair)) return pair;
      const reg = linearRegression(pair.ys, pair.xs);
      return isError(reg) ? reg : finite(reg.intercept);
    },
  },
  {
    name: 'FORECAST.LINEAR',
    category: 'statistical',
    minArgs: 3,
    maxArgs: 3,
    signature: [['x', 'known_ys', 'known_xs']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 3) return error('#VALUE!');
      const xVal = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(xVal)) return xVal;
      const x = ctx.number(xVal);
      if (!Number.isFinite(x)) return error('#VALUE!');
      const pair = collectPairs(args[1]!, args[2]!, ctx);
      if (isError(pair)) return pair;
      const reg = linearRegression(pair.ys, pair.xs);
      return isError(reg) ? reg : finite(reg.intercept + reg.slope * x);
    },
  },
];
