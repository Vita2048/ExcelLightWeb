import type { FormulaExpression } from '../../core/formula-references';
import { arrayValues, error, finite, isArray, isError, isOmitted, isRange, type FormulaError, type Scalar } from '../formula-values';
import { collectNumbers, evaluateScalarArgs } from './helpers';
import type { FunctionContext, FunctionSpec } from './types';

function numOrDefault(expr: FormulaExpression | undefined, fallback: number, ctx: FunctionContext): number | FormulaError {
  if (!expr) return fallback;
  const val = ctx.evaluate(expr);
  if (isOmitted(val)) return fallback;
  const s = ctx.scalar(val);
  if (isError(s)) return s;
  const n = ctx.number(s);
  return Number.isNaN(n) ? error('#VALUE!') : Number.isFinite(n) ? n : error('#NUM!');
}

function pmtCalc(rate: number, nper: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return -(pv + fv) / nper;
  const pow = (1 + rate) ** nper;
  return -((fv + pv * pow) * rate) / ((1 + rate * type) * (pow - 1));
}

function fvCalc(rate: number, nper: number, pmt: number, pv: number, type: number): number {
  if (rate === 0) return -(pv + pmt * nper);
  const pow = (1 + rate) ** nper;
  return -(pv * pow + (pmt * (1 + rate * type) * (pow - 1)) / rate);
}

function tvmValue(rate: number, nper: number, pmt: number, pv: number, fv: number, type: number): number {
  if (Math.abs(rate) < 1e-12) return pv + pmt * nper + fv;
  const pow = (1 + rate) ** nper;
  return pv * pow + (pmt * (1 + rate * type) * (pow - 1)) / rate + fv;
}

function collectSeries(expr: FormulaExpression, ctx: FunctionContext): number[] | FormulaError {
  const val = ctx.evaluate(expr, true);
  if (isError(val)) return val;
  const nums: number[] = [];
  const push = (s: Scalar): FormulaError | undefined => {
    if (isError(s)) return s;
    if (typeof s === 'number') nums.push(s);
    return undefined;
  };
  if (isRange(val)) {
    for (const r of val) for (const item of ctx.values(r)) {
      const err = push(item);
      if (err) return err;
    }
  } else if (isArray(val)) {
    for (const item of arrayValues(val)) {
      const err = push(item);
      if (err) return err;
    }
  } else {
    const err = push(ctx.scalar(val));
    if (err) return err;
  }
  return nums;
}

function checkCashFlows(flows: readonly number[]): boolean {
  return flows.some(v => v > 0) && flows.some(v => v < 0);
}

export const financialFunctions: readonly FunctionSpec[] = [
  {
    name: 'PMT',
    category: 'financial',
    minArgs: 3,
    maxArgs: 5,
    signature: [['rate', 'nper', 'pv', '[fv]', '[type]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 3 || args.length > 5) return error('#VALUE!');
      const rate = numOrDefault(args[0], 0, ctx);
      const nper = numOrDefault(args[1], 0, ctx);
      const pv = numOrDefault(args[2], 0, ctx);
      const fv = numOrDefault(args[3], 0, ctx);
      const type = numOrDefault(args[4], 0, ctx);
      if (isError(rate)) return rate;
      if (isError(nper)) return nper;
      if (isError(pv)) return pv;
      if (isError(fv)) return fv;
      if (isError(type)) return type;
      const t = type !== 0 ? 1 : 0;
      if (nper === 0 || (rate <= -1 && rate !== 0)) return error('#NUM!');
      return finite(pmtCalc(rate, nper, pv, fv, t));
    },
  },
  {
    name: 'PV',
    category: 'financial',
    minArgs: 3,
    maxArgs: 5,
    signature: [['rate', 'nper', 'pmt', '[fv]', '[type]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 3 || args.length > 5) return error('#VALUE!');
      const rate = numOrDefault(args[0], 0, ctx);
      const nper = numOrDefault(args[1], 0, ctx);
      const pmt = numOrDefault(args[2], 0, ctx);
      const fv = numOrDefault(args[3], 0, ctx);
      const type = numOrDefault(args[4], 0, ctx);
      if (isError(rate)) return rate;
      if (isError(nper)) return nper;
      if (isError(pmt)) return pmt;
      if (isError(fv)) return fv;
      if (isError(type)) return type;
      const t = type !== 0 ? 1 : 0;
      if (rate === -1) return error('#DIV/0!');
      if (rate === 0) return finite(-(fv + pmt * nper));
      const pow = (1 + rate) ** nper;
      if (!Number.isFinite(pow) || pow === 0) return error('#NUM!');
      return finite(-(fv + (pmt * (1 + rate * t) * (pow - 1)) / rate) / pow);
    },
  },
  {
    name: 'FV',
    category: 'financial',
    minArgs: 3,
    maxArgs: 5,
    signature: [['rate', 'nper', 'pmt', '[pv]', '[type]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 3 || args.length > 5) return error('#VALUE!');
      const rate = numOrDefault(args[0], 0, ctx);
      const nper = numOrDefault(args[1], 0, ctx);
      const pmt = numOrDefault(args[2], 0, ctx);
      const pv = numOrDefault(args[3], 0, ctx);
      const type = numOrDefault(args[4], 0, ctx);
      if (isError(rate)) return rate;
      if (isError(nper)) return nper;
      if (isError(pmt)) return pmt;
      if (isError(pv)) return pv;
      if (isError(type)) return type;
      const t = type !== 0 ? 1 : 0;
      return finite(fvCalc(rate, nper, pmt, pv, t));
    },
  },
  {
    name: 'NPV',
    category: 'financial',
    minArgs: 2,
    maxArgs: 255,
    signature: [['rate', 'value1', '[value2]'], true],
    evaluate(args, ctx): Scalar {
      if (args.length < 2) return error('#VALUE!');
      const rate = numOrDefault(args[0], 0, ctx);
      if (isError(rate)) return rate;
      if (rate <= -1) return error('#DIV/0!');
      const flows = collectNumbers(args.slice(1), ctx);
      if (isError(flows)) return flows;
      let sum = 0;
      for (let i = 0; i < flows.length; i++) sum += flows[i]! / (1 + rate) ** (i + 1);
      return finite(sum);
    },
  },
  {
    name: 'XNPV',
    category: 'financial',
    minArgs: 3,
    maxArgs: 3,
    signature: [['rate', 'values', 'dates']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 3) return error('#VALUE!');
      const rate = numOrDefault(args[0], 0, ctx);
      if (isError(rate)) return rate;
      if (rate <= -1) return error('#NUM!');
      const vals = collectSeries(args[1]!, ctx);
      if (isError(vals)) return vals;
      const dates = collectSeries(args[2]!, ctx);
      if (isError(dates)) return dates;
      if (vals.length !== dates.length || !vals.length) return error('#NUM!');
      const d0 = dates[0]!;
      let sum = 0;
      for (let i = 0; i < vals.length; i++) {
        if (dates[i]! < d0) return error('#NUM!');
        sum += vals[i]! / (1 + rate) ** ((dates[i]! - d0) / 365);
      }
      return finite(sum);
    },
  },
  {
    name: 'IRR',
    category: 'financial',
    minArgs: 1,
    maxArgs: 2,
    signature: [['values', '[guess]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 2) return error('#VALUE!');
      const flows = collectSeries(args[0]!, ctx);
      if (isError(flows)) return flows;
      const guess = numOrDefault(args[1], 0.1, ctx);
      if (isError(guess)) return guess;
      if (!checkCashFlows(flows)) return error('#NUM!');

      let r = guess;
      for (let iter = 0; iter < 100; iter++) {
        if (r <= -1) return error('#NUM!');
        let f = 0, df = 0;
        for (let i = 0; i < flows.length; i++) {
          const factor = (1 + r) ** i;
          f += flows[i]! / factor;
          if (i > 0) df -= (i * flows[i]!) / ((1 + r) ** (i + 1));
        }
        if (Math.abs(f) < 1e-12) return finite(r);
        if (!Number.isFinite(df) || df === 0) return error('#NUM!');
        const next = r - f / df;
        if (Math.abs(next - r) < 1e-12) return finite(next);
        r = next;
      }
      return error('#NUM!');
    },
  },
  {
    name: 'XIRR',
    category: 'financial',
    minArgs: 2,
    maxArgs: 3,
    signature: [['values', 'dates', '[guess]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const flows = collectSeries(args[0]!, ctx);
      if (isError(flows)) return flows;
      const dates = collectSeries(args[1]!, ctx);
      if (isError(dates)) return dates;
      const guess = numOrDefault(args[2], 0.1, ctx);
      if (isError(guess)) return guess;
      if (flows.length !== dates.length || !checkCashFlows(flows)) return error('#NUM!');
      const d0 = dates[0]!;
      for (const d of dates) if (d < d0) return error('#NUM!');

      let r = guess;
      for (let iter = 0; iter < 100; iter++) {
        if (r <= -1) return error('#NUM!');
        let f = 0, df = 0;
        for (let i = 0; i < flows.length; i++) {
          const exp = (dates[i]! - d0) / 365;
          f += flows[i]! / (1 + r) ** exp;
          if (exp !== 0) df -= (exp * flows[i]!) / (1 + r) ** (exp + 1);
        }
        if (Math.abs(f) < 1e-12) return finite(r);
        if (!Number.isFinite(df) || df === 0) return error('#NUM!');
        const next = r - f / df;
        if (Math.abs(next - r) < 1e-12) return finite(next);
        r = next;
      }
      return error('#NUM!');
    },
  },
  {
    name: 'RATE',
    category: 'financial',
    minArgs: 3,
    maxArgs: 6,
    signature: [['nper', 'pmt', 'pv', '[fv]', '[type]', '[guess]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 3 || args.length > 6) return error('#VALUE!');
      const nper = numOrDefault(args[0], 0, ctx);
      const pmt = numOrDefault(args[1], 0, ctx);
      const pv = numOrDefault(args[2], 0, ctx);
      const fv = numOrDefault(args[3], 0, ctx);
      const type = numOrDefault(args[4], 0, ctx);
      const guess = numOrDefault(args[5], 0.1, ctx);
      if (isError(nper)) return nper;
      if (isError(pmt)) return pmt;
      if (isError(pv)) return pv;
      if (isError(fv)) return fv;
      if (isError(type)) return type;
      if (isError(guess)) return guess;
      if (nper <= 0) return error('#NUM!');
      const t = type !== 0 ? 1 : 0;

      let r = guess;
      for (let iter = 0; iter < 100; iter++) {
        if (r <= -1) return error('#NUM!');
        const f = tvmValue(r, nper, pmt, pv, fv, t);
        if (Math.abs(f) < 1e-7) return finite(r);
        const h = Math.max(1e-6, Math.abs(r) * 1e-6);
        const df = (tvmValue(r + h, nper, pmt, pv, fv, t) - f) / h;
        if (!Number.isFinite(df) || df === 0) return error('#NUM!');
        const next = r - f / df;
        if (Math.abs(next - r) < 1e-7) return finite(next);
        r = next;
      }
      return error('#NUM!');
    },
  },
  {
    name: 'NPER',
    category: 'financial',
    minArgs: 3,
    maxArgs: 5,
    signature: [['rate', 'pmt', 'pv', '[fv]', '[type]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 3 || args.length > 5) return error('#VALUE!');
      const rate = numOrDefault(args[0], 0, ctx);
      const pmt = numOrDefault(args[1], 0, ctx);
      const pv = numOrDefault(args[2], 0, ctx);
      const fv = numOrDefault(args[3], 0, ctx);
      const type = numOrDefault(args[4], 0, ctx);
      if (isError(rate)) return rate;
      if (isError(pmt)) return pmt;
      if (isError(pv)) return pv;
      if (isError(fv)) return fv;
      if (isError(type)) return type;
      const t = type !== 0 ? 1 : 0;
      if (rate <= -1) return error('#NUM!');
      if (rate === 0) return pmt === 0 ? error('#DIV/0!') : finite(-(pv + fv) / pmt);
      const z = (pmt * (1 + rate * t)) / rate;
      const ratio = (z - fv) / (z + pv);
      if (ratio <= 0) return error('#NUM!');
      return finite(Math.log(ratio) / Math.log(1 + rate));
    },
  },
  {
    name: 'IPMT',
    category: 'financial',
    minArgs: 4,
    maxArgs: 6,
    signature: [['rate', 'per', 'nper', 'pv', '[fv]', '[type]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 4 || args.length > 6) return error('#VALUE!');
      const rate = numOrDefault(args[0], 0, ctx);
      const per = numOrDefault(args[1], 0, ctx);
      const nper = numOrDefault(args[2], 0, ctx);
      const pv = numOrDefault(args[3], 0, ctx);
      const fv = numOrDefault(args[4], 0, ctx);
      const type = numOrDefault(args[5], 0, ctx);
      if (isError(rate)) return rate;
      if (isError(per)) return per;
      if (isError(nper)) return nper;
      if (isError(pv)) return pv;
      if (isError(fv)) return fv;
      if (isError(type)) return type;
      const t = type !== 0 ? 1 : 0;
      if (per < 1 || per > nper || nper <= 0 || rate <= -1) return error('#NUM!');
      if (rate === 0 || (t === 1 && per === 1)) return 0;
      const pmt = pmtCalc(rate, nper, pv, fv, t);
      const rawIpmt = fvCalc(rate, per - 1, pmt, pv, t) * rate;
      return finite(t === 1 ? rawIpmt / (1 + rate) : rawIpmt);
    },
  },
  {
    name: 'PPMT',
    category: 'financial',
    minArgs: 4,
    maxArgs: 6,
    signature: [['rate', 'per', 'nper', 'pv', '[fv]', '[type]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 4 || args.length > 6) return error('#VALUE!');
      const rate = numOrDefault(args[0], 0, ctx);
      const per = numOrDefault(args[1], 0, ctx);
      const nper = numOrDefault(args[2], 0, ctx);
      const pv = numOrDefault(args[3], 0, ctx);
      const fv = numOrDefault(args[4], 0, ctx);
      const type = numOrDefault(args[5], 0, ctx);
      if (isError(rate)) return rate;
      if (isError(per)) return per;
      if (isError(nper)) return nper;
      if (isError(pv)) return pv;
      if (isError(fv)) return fv;
      if (isError(type)) return type;
      const t = type !== 0 ? 1 : 0;
      if (per < 1 || per > nper || nper <= 0 || rate <= -1) return error('#NUM!');
      const pmt = pmtCalc(rate, nper, pv, fv, t);
      if (rate === 0 || (t === 1 && per === 1)) return finite(pmt);
      const rawIpmt = fvCalc(rate, per - 1, pmt, pv, t) * rate;
      const ipmt = t === 1 ? rawIpmt / (1 + rate) : rawIpmt;
      return finite(pmt - ipmt);
    },
  },
  {
    name: 'SLN',
    category: 'financial',
    minArgs: 3,
    maxArgs: 3,
    signature: [['cost', 'salvage', 'life']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 3) return error('#VALUE!');
      const cost = ctx.number(values[0]!);
      const salvage = ctx.number(values[1]!);
      const life = ctx.number(values[2]!);
      if (Number.isNaN(cost) || Number.isNaN(salvage) || Number.isNaN(life)) return error('#VALUE!');
      if (life === 0) return error('#DIV/0!');
      return finite((cost - salvage) / life);
    },
  },
  {
    name: 'DB',
    category: 'financial',
    minArgs: 4,
    maxArgs: 5,
    signature: [['cost', 'salvage', 'life', 'period', '[month]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 4 || args.length > 5) return error('#VALUE!');
      const cost = numOrDefault(args[0], 0, ctx);
      const salvage = numOrDefault(args[1], 0, ctx);
      const life = numOrDefault(args[2], 0, ctx);
      const period = numOrDefault(args[3], 0, ctx);
      const month = numOrDefault(args[4], 12, ctx);
      if (isError(cost)) return cost;
      if (isError(salvage)) return salvage;
      if (isError(life)) return life;
      if (isError(period)) return period;
      if (isError(month)) return month;
      const m = Math.trunc(month);
      const p = Math.trunc(period);
      const l = Math.trunc(life);
      if (cost < 0 || salvage < 0 || l < 1 || m < 1 || m > 12 || p < 1 || p > (m === 12 ? l : l + 1)) {
        return error('#NUM!');
      }
      if (cost === 0) return 0;
      const rate = Math.round((1 - (salvage / cost) ** (1 / l)) * 1000) / 1000;
      let totalDepr = 0;
      let currentDepr = 0;
      for (let i = 1; i <= p; i++) {
        if (i === 1) currentDepr = cost * rate * (m / 12);
        else if (i === l + 1) currentDepr = (cost - totalDepr) * rate * ((12 - m) / 12);
        else currentDepr = (cost - totalDepr) * rate;
        totalDepr += currentDepr;
      }
      return finite(currentDepr);
    },
  },
  {
    name: 'DDB',
    category: 'financial',
    minArgs: 4,
    maxArgs: 5,
    signature: [['cost', 'salvage', 'life', 'period', '[factor]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 4 || args.length > 5) return error('#VALUE!');
      const cost = numOrDefault(args[0], 0, ctx);
      const salvage = numOrDefault(args[1], 0, ctx);
      const life = numOrDefault(args[2], 0, ctx);
      const period = numOrDefault(args[3], 0, ctx);
      const factor = numOrDefault(args[4], 2, ctx);
      if (isError(cost)) return cost;
      if (isError(salvage)) return salvage;
      if (isError(life)) return life;
      if (isError(period)) return period;
      if (isError(factor)) return factor;
      if (cost < 0 || salvage < 0 || life <= 0 || period < 1 || period > life || factor <= 0) {
        return error('#NUM!');
      }
      if (cost <= salvage) return 0;
      let book = cost;
      let depr = 0;
      const wholePeriods = Math.floor(period);
      const frac = period - wholePeriods;
      for (let i = 1; i <= wholePeriods; i++) {
        depr = Math.min((book * factor) / life, Math.max(0, book - salvage));
        book -= depr;
      }
      if (frac > 0) {
        depr = Math.min((book * factor * frac) / life, Math.max(0, book - salvage));
      }
      return finite(depr);
    },
  },
];
