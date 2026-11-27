import type { FormulaExpression } from '../../core/formula-references';
import {
  arrayFromMatrix, error, finite, isArray, isError, isOmitted, isRange,
  type FormulaArray, type FormulaError, type FormulaValue, type Range, type Scalar,
} from '../formula-values';
import {
  aggregate, aggregateInputs, collectNumbers, conditional, evaluateScalarArgs, logicalValue,
  medianOf, modeOf, multiConditional, nthExtreme, percentileOf,
} from './helpers';
import type { FunctionContext, FunctionSignature, FunctionSpec } from './types';

const rounding: FunctionSignature = [['number', 'num_digits']];
const singleNum: FunctionSignature = [['number']];

function round(name: string, number: number, places: number): Scalar {
  places = Math.trunc(places);
  if (name !== 'ROUND' && (places > 308 || (name === 'ROUNDUP' && places < -308))) return error('#VALUE!');
  if (number === 0 || places > 308) return number;
  if (places < -308) return 0;
  function shift(value: number, amount: number) {
    const [digits, exponent] = String(value).split('e');
    return Number(digits + 'e' + (Number(exponent ?? 0) + amount));
  }
  const scaled = shift(Math.abs(number), places);
  if (!Number.isFinite(scaled)) return number;
  const operation = name === 'ROUNDUP' ? Math.ceil : name === 'ROUNDDOWN' ? Math.floor : Math.round;
  return finite(Math.sign(number) * shift(operation(scaled), -places));
}

function evaluateRound(name: 'ROUND' | 'ROUNDUP' | 'ROUNDDOWN', args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
  const values = evaluateScalarArgs(args, ctx);
  if (!Array.isArray(values)) return values;
  if (values.length !== 2) return error('#VALUE!');
  const number = ctx.number(values[0] ?? null), places = ctx.number(values[1]!);
  return Number.isNaN(number) || Number.isNaN(places) ? error('#VALUE!') :
    Number.isFinite(number) && Number.isFinite(places) ? round(name, number, places) : error('#NUM!');
}

function oneNum(
  name: string,
  fn: (value: number) => Scalar,
  signature: FunctionSignature = singleNum,
): FunctionSpec {
  return {
    name,
    category: 'math',
    minArgs: 1,
    maxArgs: 1,
    signature,
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const n = ctx.number(values[0] ?? null);
      return Number.isNaN(n) ? error('#VALUE!') : Number.isFinite(n) ? fn(n) : error('#NUM!');
    },
  };
}

function twoNums(
  name: string,
  signature: FunctionSignature,
  fn: (a: number, b: number) => Scalar,
): FunctionSpec {
  return {
    name,
    category: 'math',
    minArgs: 2,
    maxArgs: 2,
    signature,
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 2) return error('#VALUE!');
      const a = ctx.number(values[0] ?? null);
      const b = ctx.number(values[1] ?? null);
      return Number.isNaN(a) || Number.isNaN(b) ? error('#VALUE!') :
        Number.isFinite(a) && Number.isFinite(b) ? fn(a, b) : error('#NUM!');
    },
  };
}

function numArgOrDefault(expr: FormulaExpression | undefined, fallback: number, ctx: FunctionContext): number | FormulaError {
  if (!expr) return fallback;
  const raw = ctx.evaluate(expr);
  if (isOmitted(raw)) return fallback;
  const s = ctx.scalar(raw);
  if (isError(s)) return s;
  const n = ctx.number(s);
  return Number.isNaN(n) ? error('#VALUE!') : Number.isFinite(n) ? n : error('#NUM!');
}

function gcdPair(a: number, b: number): number {
  let x = Math.abs(Math.trunc(a));
  let y = Math.abs(Math.trunc(b));
  while (y !== 0) {
    const t = x % y;
    x = y;
    y = t;
  }
  return x;
}

function collectWholeInts(args: readonly FormulaExpression[], ctx: FunctionContext): number[] | FormulaError {
  const out: number[] = [];
  for (const val of aggregateInputs(args, ctx, 2)) {
    if (isError(val)) return val;
    if (typeof val !== 'number') continue;
    if (val < 0) return error('#NUM!');
    out.push(Math.trunc(val));
  }
  return out;
}

function subtotal(args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
  if (args.length < 2 || args.length > 255) return error('#VALUE!');
  const value = ctx.scalar(ctx.evaluate(args[0]!));
  if (isError(value)) return value;
  if (typeof value !== 'number') return error('#VALUE!');
  const code = Math.trunc(value), operation = code > 100 ? code - 100 : code;
  if (operation < 1 || operation > 11) return error('#VALUE!');
  const ranges: Range[] = [];
  for (const arg of args.slice(1)) {
    if (arg.kind !== 'reference') return error('#VALUE!');
    const target = ctx.evaluate(arg);
    if (!isRange(target) || target.length !== 1) return error('#VALUE!');
    ranges.push(target[0]!);
  }
  function* eligible() {
    for (const range of ranges) yield* ctx.values(range, { skipHidden: code > 100 });
  }
  return aggregate(operation, eligible());
}

function evaluateAggregate(args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
  if (args.length < 3 || args.length > 255) return error('#VALUE!');
  const codeVal = ctx.scalar(ctx.evaluate(args[0]!));
  if (isError(codeVal)) return codeVal;
  const code = Math.trunc(ctx.number(codeVal));
  const optVal = ctx.scalar(ctx.evaluate(args[1]!));
  if (isError(optVal)) return optVal;
  const opt = Math.trunc(ctx.number(optVal));
  if (!Number.isFinite(code) || code < 1 || code > 19 || !Number.isFinite(opt) || opt < 0 || opt > 7) {
    return error('#VALUE!');
  }

  const ignoreNested = opt <= 3;
  const skipHidden = opt === 1 || opt === 3 || opt === 5 || opt === 7;
  const ignoreErrors = opt === 2 || opt === 3 || opt === 6 || opt === 7;

  if (code >= 12) {
    const srcVal = ctx.evaluate(args[2]!, true);
    if (isError(srcVal)) return srcVal;
    const nums: number[] = [];
    const pushItem = (item: Scalar): FormulaError | undefined => {
      if (isError(item)) return ignoreErrors ? undefined : item;
      if (typeof item === 'number') nums.push(item);
      return undefined;
    };
    if (isRange(srcVal)) {
      for (const r of srcVal) {
        const iter = ignoreNested ? ctx.values(r, { skipHidden }) : ctx.values(r);
        for (const item of iter) {
          const err = pushItem(item);
          if (err) return err;
        }
      }
    } else if (isArray(srcVal)) {
      for (let r = 0; r < srcVal.rows; r++) for (let c = 0; c < srcVal.columns; c++) {
        const err = pushItem(srcVal.at(r, c));
        if (err) return err;
      }
    } else {
      return error('#VALUE!');
    }

    if (code === 12) return medianOf(nums);
    if (code === 13) return modeOf(nums);
    if (args.length < 4) return error('#VALUE!');
    const kVal = ctx.scalar(ctx.evaluate(args[3]!));
    if (isError(kVal)) return kVal;
    const k = ctx.number(kVal);
    if (!Number.isFinite(k)) return error('#VALUE!');
    if (code === 14 || code === 15) return nthExtreme(nums, k, code === 14);
    if (code === 16 || code === 18) return percentileOf(nums, k, code === 16);
    const q = Math.trunc(k);
    if (code === 17 && (q < 0 || q > 4)) return error('#NUM!');
    if (code === 19 && (q < 1 || q > 3)) return error('#NUM!');
    return percentileOf(nums, q / 4, code === 17);
  }

  const collected: Scalar[] = [];
  for (const arg of args.slice(2)) {
    const val = ctx.evaluate(arg, true);
    if (isError(val)) {
      if (!ignoreErrors) return val;
      continue;
    }
    if (isRange(val)) {
      for (const r of val) {
        const iter = ignoreNested ? ctx.values(r, { skipHidden }) : ctx.values(r);
        for (const item of iter) {
          if (ignoreErrors && isError(item)) continue;
          collected.push(item);
        }
      }
    } else if (isArray(val)) {
      for (let r = 0; r < val.rows; r++) for (let c = 0; c < val.columns; c++) {
        const item = val.at(r, c);
        if (ignoreErrors && isError(item)) continue;
        collected.push(item);
      }
    } else {
      const s = ctx.scalar(val);
      if (ignoreErrors && isError(s)) continue;
      collected.push(s);
    }
  }
  return aggregate(code, collected);
}

function numericMatrix(arr: FormulaArray): number[] | FormulaError {
  const out = new Array<number>(arr.rows * arr.columns);
  for (let r = 0; r < arr.rows; r++) for (let c = 0; c < arr.columns; c++) {
    const v = arr.at(r, c);
    if (isError(v)) return v;
    if (typeof v !== 'number') return error('#VALUE!');
    out[r * arr.columns + c] = v;
  }
  return out;
}

export const mathFunctions: readonly FunctionSpec[] = [
  {
    name: 'SUM',
    category: 'math',
    minArgs: 1,
    maxArgs: 255,
    signature: [['number1', '[number2]'], true],
    evaluate(args, ctx) {
      if (!args.length) return error('#VALUE!');
      return aggregate(9, aggregateInputs(args, ctx, 9));
    },
  },
  {
    name: 'PRODUCT',
    category: 'math',
    minArgs: 1,
    maxArgs: 255,
    signature: [['number1', '[number2]'], true],
    evaluate(args, ctx) {
      if (!args.length) return error('#VALUE!');
      return aggregate(6, aggregateInputs(args, ctx, 6));
    },
  },
  {
    name: 'SUMSQ',
    category: 'math',
    minArgs: 1,
    maxArgs: 255,
    signature: [['number1', '[number2]'], true],
    evaluate(args, ctx) {
      const nums = collectNumbers(args, ctx);
      if (isError(nums)) return nums;
      let total = 0;
      for (const n of nums) total += n * n;
      return finite(total);
    },
  },
  {
    name: 'SUMPRODUCT',
    category: 'math',
    minArgs: 1,
    maxArgs: 255,
    signature: [['array1', '[array2]'], true],
    evaluate(args, ctx): Scalar {
      if (!args.length) return error('#VALUE!');
      const arrays: FormulaArray[] = [];
      let rows = 0;
      let cols = 0;
      for (const arg of args) {
        const val = ctx.evaluate(arg, true);
        if (isError(val)) return val;
        const arr = ctx.arrayOf(val);
        if (!arr) return error('#VALUE!');
        if (!arrays.length) {
          rows = arr.rows;
          cols = arr.columns;
        } else if (arr.rows !== rows || arr.columns !== cols) {
          return error('#VALUE!');
        }
        arrays.push(arr);
      }
      let sum = 0;
      for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
        let prod = 1;
        for (const arr of arrays) {
          const v = arr.at(r, c);
          if (isError(v)) return v;
          prod *= typeof v === 'number' ? v : 0;
        }
        sum += prod;
      }
      return finite(sum);
    },
  },
  {
    name: 'ABS',
    category: 'math',
    minArgs: 1,
    maxArgs: 1,
    signature: singleNum,
    evaluate(args, ctx) {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const number = ctx.number(values[0] ?? null);
      return Number.isNaN(number) ? error('#VALUE!') : finite(Math.abs(number));
    },
  },
  {
    name: 'ROUND',
    category: 'math',
    minArgs: 2,
    maxArgs: 2,
    signature: rounding,
    evaluate: (args, ctx) => evaluateRound('ROUND', args, ctx),
  },
  {
    name: 'ROUNDUP',
    category: 'math',
    minArgs: 2,
    maxArgs: 2,
    signature: rounding,
    evaluate: (args, ctx) => evaluateRound('ROUNDUP', args, ctx),
  },
  {
    name: 'ROUNDDOWN',
    category: 'math',
    minArgs: 2,
    maxArgs: 2,
    signature: rounding,
    evaluate: (args, ctx) => evaluateRound('ROUNDDOWN', args, ctx),
  },
  {
    name: 'INT',
    category: 'math',
    minArgs: 1,
    maxArgs: 1,
    signature: singleNum,
    evaluate: (args, ctx) => oneNum('INT', n => finite(Math.floor(n))).evaluate(args, ctx),
  },
  {
    name: 'TRUNC',
    category: 'math',
    minArgs: 1,
    maxArgs: 2,
    signature: [['number', '[num_digits]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 2) return error('#VALUE!');
      const n = numArgOrDefault(args[0], 0, ctx);
      if (isError(n)) return n;
      const d = numArgOrDefault(args[1], 0, ctx);
      if (isError(d)) return d;
      return round('ROUNDDOWN', n, d);
    },
  },
  twoNums('MOD', [['number', 'divisor']], (n, d) => d === 0 ? error('#DIV/0!') : finite(n - d * Math.floor(n / d))),
  twoNums('POWER', [['number', 'power']], (base, exp) => {
    if (base === 0 && exp === 0) return error('#NUM!');
    if (base === 0 && exp < 0) return error('#DIV/0!');
    const res = Math.pow(base, exp);
    return Number.isNaN(res) ? error('#NUM!') : finite(res);
  }),
  oneNum('SQRT', n => n < 0 ? error('#NUM!') : finite(Math.sqrt(n))),
  oneNum('SIGN', n => n > 0 ? 1 : n < 0 ? -1 : 0),
  {
    name: 'RAND',
    category: 'math',
    minArgs: 0,
    maxArgs: 0,
    volatile: true,
    signature: [[]],
    evaluate: (args, ctx) => args.length === 0 ? finite(ctx.random()) : error('#VALUE!'),
  },
  {
    name: 'RANDBETWEEN',
    category: 'math',
    minArgs: 2,
    maxArgs: 2,
    volatile: true,
    signature: [['bottom', 'top']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 2) return error('#VALUE!');
      const a = ctx.number(values[0] ?? null);
      const b = ctx.number(values[1] ?? null);
      if (Number.isNaN(a) || Number.isNaN(b)) return error('#VALUE!');
      if (!Number.isFinite(a) || !Number.isFinite(b) || a > b) return error('#NUM!');
      const lo = Math.ceil(a);
      const hi = Math.floor(b);
      if (lo > hi) return error('#NUM!');
      const u = Math.min(Math.max(ctx.random(), 0), 1 - Number.EPSILON);
      return lo + Math.floor(u * (hi - lo + 1));
    },
  },
  twoNums('CEILING', [['number', 'significance']], (n, s) => {
    if (s === 0) return 0;
    if (n > 0 && s < 0) return error('#NUM!');
    if (n < 0 && s < 0) return finite(-Math.ceil(Math.abs(n) / Math.abs(s)) * Math.abs(s));
    return finite(Math.ceil(n / s) * s);
  }),
  {
    name: 'CEILING.MATH',
    category: 'math',
    minArgs: 1,
    maxArgs: 3,
    signature: [['number', '[significance]', '[mode]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 3) return error('#VALUE!');
      const n = numArgOrDefault(args[0], 0, ctx);
      if (isError(n)) return n;
      const sig = numArgOrDefault(args[1], 1, ctx);
      if (isError(sig)) return sig;
      const mode = numArgOrDefault(args[2], 0, ctx);
      if (isError(mode)) return mode;
      if (sig === 0 || n === 0) return 0;
      const s = Math.abs(sig);
      if (n > 0) return finite(Math.ceil(n / s) * s);
      return mode !== 0 ? finite(-Math.ceil(Math.abs(n) / s) * s) : finite(-Math.floor(Math.abs(n) / s) * s);
    },
  },
  twoNums('FLOOR', [['number', 'significance']], (n, s) => {
    if (s === 0) return error('#DIV/0!');
    if (n > 0 && s < 0) return error('#NUM!');
    if (n < 0 && s < 0) return finite(-Math.floor(Math.abs(n) / Math.abs(s)) * Math.abs(s));
    return finite(Math.floor(n / s) * s);
  }),
  {
    name: 'FLOOR.MATH',
    category: 'math',
    minArgs: 1,
    maxArgs: 3,
    signature: [['number', '[significance]', '[mode]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 3) return error('#VALUE!');
      const n = numArgOrDefault(args[0], 0, ctx);
      if (isError(n)) return n;
      const sig = numArgOrDefault(args[1], 1, ctx);
      if (isError(sig)) return sig;
      const mode = numArgOrDefault(args[2], 0, ctx);
      if (isError(mode)) return mode;
      if (sig === 0 || n === 0) return 0;
      const s = Math.abs(sig);
      if (n > 0) return finite(Math.floor(n / s) * s);
      return mode !== 0 ? finite(-Math.floor(Math.abs(n) / s) * s) : finite(-Math.ceil(Math.abs(n) / s) * s);
    },
  },
  twoNums('MROUND', [['number', 'multiple']], (n, m) => {
    if (m === 0) return 0;
    if ((n > 0 && m < 0) || (n < 0 && m > 0)) return error('#NUM!');
    return finite(Math.round(n / m) * m);
  }),
  oneNum('EXP', n => finite(Math.exp(n))),
  oneNum('LN', n => n <= 0 ? error('#NUM!') : finite(Math.log(n))),
  {
    name: 'LOG',
    category: 'math',
    minArgs: 1,
    maxArgs: 2,
    signature: [['number', '[base]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 2) return error('#VALUE!');
      const x = numArgOrDefault(args[0], 0, ctx);
      if (isError(x)) return x;
      const base = numArgOrDefault(args[1], 10, ctx);
      if (isError(base)) return base;
      if (x <= 0 || base <= 0) return error('#NUM!');
      if (base === 1) return error('#DIV/0!');
      return finite(Math.log(x) / Math.log(base));
    },
  },
  oneNum('LOG10', n => n <= 0 ? error('#NUM!') : finite(Math.log10(n))),
  {
    name: 'PI',
    category: 'math',
    minArgs: 0,
    maxArgs: 0,
    signature: [[]],
    evaluate: args => args.length === 0 ? Math.PI : error('#VALUE!'),
  },
  oneNum('SIN', n => finite(Math.sin(n))),
  oneNum('COS', n => finite(Math.cos(n))),
  oneNum('TAN', n => finite(Math.tan(n))),
  oneNum('ASIN', n => Math.abs(n) > 1 ? error('#NUM!') : finite(Math.asin(n))),
  oneNum('ACOS', n => Math.abs(n) > 1 ? error('#NUM!') : finite(Math.acos(n))),
  oneNum('ATAN', n => finite(Math.atan(n))),
  twoNums('ATAN2', [['x_num', 'y_num']], (x, y) => x === 0 && y === 0 ? error('#DIV/0!') : finite(Math.atan2(y, x))),
  oneNum('DEGREES', n => finite((n * 180) / Math.PI), [['angle']]),
  oneNum('RADIANS', n => finite((n * Math.PI) / 180), [['angle']]),
  oneNum('FACT', n => {
    if (n < 0) return error('#NUM!');
    const k = Math.trunc(n);
    let prod = 1;
    for (let i = 2; i <= k; i++) {
      prod *= i;
      if (!Number.isFinite(prod)) return error('#NUM!');
    }
    return prod;
  }),
  twoNums('COMBIN', [['number', 'number_chosen']], (rawN, rawK) => {
    const n = Math.trunc(rawN);
    const k = Math.trunc(rawK);
    if (n < 0 || k < 0 || k > n) return error('#NUM!');
    const r = Math.min(k, n - k);
    let res = 1;
    for (let i = 1; i <= r; i++) {
      res = (res * (n - r + i)) / i;
      if (!Number.isFinite(res)) return error('#NUM!');
    }
    return Math.round(res);
  }),
  twoNums('PERMUT', [['number', 'number_chosen']], (rawN, rawK) => {
    const n = Math.trunc(rawN);
    const k = Math.trunc(rawK);
    if (n < 0 || k < 0 || k > n) return error('#NUM!');
    let res = 1;
    for (let i = 0; i < k; i++) {
      res *= n - i;
      if (!Number.isFinite(res)) return error('#NUM!');
    }
    return res;
  }),
  {
    name: 'GCD',
    category: 'math',
    minArgs: 1,
    maxArgs: 255,
    signature: [['number1', '[number2]'], true],
    evaluate(args, ctx): Scalar {
      const ints = collectWholeInts(args, ctx);
      if (isError(ints)) return ints;
      if (!ints.length) return 0;
      let g = ints[0]!;
      for (let i = 1; i < ints.length; i++) g = gcdPair(g, ints[i]!);
      return g;
    },
  },
  {
    name: 'LCM',
    category: 'math',
    minArgs: 1,
    maxArgs: 255,
    signature: [['number1', '[number2]'], true],
    evaluate(args, ctx): Scalar {
      const ints = collectWholeInts(args, ctx);
      if (isError(ints)) return ints;
      if (!ints.length) return 0;
      let res = ints[0]!;
      for (let i = 1; i < ints.length; i++) {
        const v = ints[i]!;
        if (res === 0 || v === 0) return 0;
        res = (res / gcdPair(res, v)) * v;
        if (!Number.isFinite(res)) return error('#NUM!');
      }
      return res;
    },
  },
  twoNums('QUOTIENT', [['numerator', 'denominator']], (n, d) => d === 0 ? error('#DIV/0!') : finite(Math.trunc(n / d))),
  oneNum('EVEN', n => {
    if (n === 0) return 0;
    const mag = Math.ceil(Math.abs(n) / 2) * 2;
    return finite(n < 0 ? -mag : mag);
  }),
  oneNum('ODD', n => {
    if (n === 0) return 1;
    let mag = Math.ceil(Math.abs(n));
    if (mag % 2 === 0) mag += 1;
    return finite(n < 0 ? -mag : mag);
  }),
  {
    name: 'MMULT',
    category: 'math',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array1', 'array2']],
    evaluate(args, ctx): FormulaValue {
      if (args.length !== 2) return error('#VALUE!');
      const v1 = ctx.evaluate(args[0]!, true);
      const v2 = ctx.evaluate(args[1]!, true);
      if (isError(v1)) return v1;
      if (isError(v2)) return v2;
      const a = ctx.arrayOf(v1);
      const b = ctx.arrayOf(v2);
      if (!a || !b || a.columns !== b.rows) return error('#VALUE!');
      const ma = numericMatrix(a);
      if (isError(ma)) return ma;
      const mb = numericMatrix(b);
      if (isError(mb)) return mb;
      const cells: Scalar[] = new Array(a.rows * b.columns);
      for (let r = 0; r < a.rows; r++) for (let c = 0; c < b.columns; c++) {
        let sum = 0;
        for (let k = 0; k < a.columns; k++) sum += ma[r * a.columns + k]! * mb[k * b.columns + c]!;
        cells[r * b.columns + c] = finite(sum);
      }
      return arrayFromMatrix(a.rows, b.columns, cells);
    },
  },
  {
    name: 'TRANSPOSE',
    category: 'math',
    minArgs: 1,
    maxArgs: 1,
    signature: [['array']],
    evaluate(args, ctx): FormulaValue {
      if (args.length !== 1) return error('#VALUE!');
      const val = ctx.evaluate(args[0]!, true);
      if (isError(val)) return val;
      const arr = ctx.arrayOf(val);
      if (!arr) return error('#VALUE!');
      const cells: Scalar[] = new Array(arr.columns * arr.rows);
      for (let r = 0; r < arr.rows; r++) for (let c = 0; c < arr.columns; c++) {
        cells[c * arr.rows + r] = arr.at(r, c);
      }
      return arrayFromMatrix(arr.columns, arr.rows, cells);
    },
  },
  {
    name: 'SUMIF',
    category: 'math',
    minArgs: 2,
    maxArgs: 3,
    signature: [['range', 'criteria', '[sum_range]']],
    evaluate: (args, ctx) => conditional('SUMIF', args, ctx),
  },
  {
    name: 'SUMIFS',
    category: 'math',
    minArgs: 3,
    maxArgs: 255,
    signature: [['sum_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]'], true],
    evaluate: (args, ctx) => multiConditional('SUMIFS', args, ctx),
  },
  {
    name: 'SUBTOTAL',
    category: 'math',
    minArgs: 2,
    maxArgs: 255,
    signature: [['function_num', 'ref1', '[ref2]'], true],
    evaluate: subtotal,
  },
  {
    name: 'AGGREGATE',
    category: 'math',
    minArgs: 3,
    maxArgs: 255,
    signature: [['function_num', 'options', 'ref1', '[ref2]'], true],
    evaluate: evaluateAggregate,
  },
];
