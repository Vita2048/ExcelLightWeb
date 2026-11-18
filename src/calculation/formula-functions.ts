import type { FormulaExpression } from '../core/formula-references';
import { documentLimits } from '../core/workbook';
import { dateFunction } from './formula-dates';
import { arrayValues, compare, error, finite, isArray, isError, isRange, matchesComparison,
  type FormulaError, type FormulaValue, type Range, type Scalar } from './formula-values';

export const nativeFunctions = [
  'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'IF', 'AND', 'OR', 'NOT', 'IFERROR',
  'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'SUMIF', 'COUNTIF', 'LEFT', 'RIGHT', 'MID',
  'LEN', 'TRIM', 'UPPER', 'LOWER', 'DATE', 'YEAR', 'MONTH', 'DAY', 'TODAY', 'VLOOKUP', 'SUBTOTAL',
] as const;
export type NativeFunction = typeof nativeFunctions[number];
export const supportedFunctions: ReadonlySet<string> = new Set(nativeFunctions);

interface FunctionContext {
  readonly evaluate: (expression: FormulaExpression, includeRanges?: boolean) => FormulaValue;
  readonly scalar: (value: FormulaValue) => Scalar;
  readonly read: (range: Range, row: number, column: number) => Scalar;
  readonly values: (range: Range, subtotal?: { skipHidden: boolean }) => Iterable<Scalar>;
  readonly entries: (range: Range) => Iterable<{ row: number; column: number; value: Scalar }>;
  readonly number: (value: Scalar) => number;
  readonly text: (value: Exclude<Scalar, FormulaError>) => string;
  readonly locale: string;
  readonly today: number;
}

function logicalValue(value: Scalar): boolean | undefined {
  if (typeof value === 'string') {
    const key = value.trim().toUpperCase();
    return key === 'TRUE' ? true : key === 'FALSE' ? false : undefined;
  }
  return isError(value) ? undefined : Boolean(value);
}

function aggregate(code: number, values: Iterable<Scalar>): Scalar {
  let count = 0, populated = 0, sum = 0, min = Infinity, max = -Infinity, product = 1;
  let mean = 0, deviation = 0;
  for (const value of values) {
    if (value !== null) populated++;
    if (isError(value) && code !== 2 && code !== 3) return value;
    if (typeof value !== 'number') continue;
    count++;
    sum += value;
    min = Math.min(min, value);
    max = Math.max(max, value);
    product *= value;
    const delta = value - mean;
    mean += delta / count;
    deviation += delta * (value - mean);
  }
  switch (code) {
    case 1: return count ? finite(sum / count) : error('#DIV/0!');
    case 2: return count;
    case 3: return populated;
    case 4: return count ? max : 0;
    case 5: return count ? min : 0;
    case 6: return count ? finite(product) : 0;
    case 9: return finite(sum);
    default: {
      const divisor = count - (code === 7 || code === 10 ? 1 : 0);
      if (divisor < 1) return error('#DIV/0!');
      const variance = Math.max(0, deviation / divisor);
      return finite(code === 7 || code === 8 ? Math.sqrt(variance) : variance);
    }
  }
}

function* aggregateInputs(args: readonly FormulaExpression[], ctx: FunctionContext, code: number): Iterable<Scalar> {
  for (const arg of args) {
    const value = ctx.evaluate(arg, true);
    if (isRange(value)) {
      for (const range of value) yield* ctx.values(range);
    } else if (isArray(value)) {
      for (const item of arrayValues(value)) {
        if (code !== 3 || typeof item === 'number' || isError(item)) yield item;
      }
    } else if (code === 3 || isError(value) || value === null) {
      yield value;
    } else {
      const numeric = ctx.number(value);
      yield Number.isNaN(numeric) ? error('#VALUE!') : finite(numeric);
    }
  }
}

const aggregateCodes = { AVERAGE: 1, COUNT: 2, COUNTA: 3, MAX: 4, MIN: 5, SUM: 9 } as const;
const comparisons = /^(<=|>=|<>|=|<|>)/u;

function wildcard(text: string): (value: string) => boolean {
  // Literal, single-character and star tokens avoid compiling workbook text as a regex.
  const tokens: { kind: 'literal' | 'one' | 'star'; value?: string }[] = [];
  const pattern = text.toUpperCase();
  for (let i = 0; i < pattern.length; i++) {
    const char = pattern[i]!;
    if (char === '~' && ['~', '?', '*'].includes(pattern[i + 1] ?? '')) {
      tokens.push({ kind: 'literal', value: pattern[++i]! });
    } else if (char === '?') tokens.push({ kind: 'one' });
    else if (char === '*') {
      if (tokens.at(-1)?.kind !== 'star') tokens.push({ kind: 'star' });
    } else tokens.push({ kind: 'literal', value: char });
  }
  return value => {
    const input = value.toUpperCase();
    let i = 0, j = 0, star = -1, retry = 0;
    while (i < input.length) {
      const token = tokens[j];
      if (token?.kind === 'one' || token?.kind === 'literal' && token.value === input[i]) { i++; j++; }
      else if (token?.kind === 'star') { star = j++; retry = i; }
      else if (star >= 0) { j = star + 1; i = ++retry; }
      else return false;
    }
    while (tokens[j]?.kind === 'star') j++;
    return j === tokens.length;
  };
}

function criteria(value: Exclude<Scalar, FormulaError>, number: FunctionContext['number']): (input: Scalar) => boolean {
  const operator = typeof value === 'string' ? comparisons.exec(value)?.[0] ?? '=' : '=';
  const raw = typeof value === 'string' && comparisons.test(value) ? value.slice(operator.length) : value;
  const boolean = typeof raw === 'string' && /^(?:TRUE|FALSE)$/iu.test(raw) ? logicalValue(raw) : undefined;
  const numeric = raw !== '' && raw !== null && typeof raw !== 'boolean' ? number(raw) : NaN;
  const predicate = typeof raw === 'string' && /[*?~]/u.test(raw) && (operator === '=' || operator === '<>')
    ? wildcard(raw) : undefined;
  return input => {
    if (isError(input)) return typeof raw === 'string' && (operator === '=' || operator === '<>') &&
      (input.value === raw.toUpperCase()) === (operator === '=');
    if (predicate) return (typeof input === 'string' && predicate(input)) === (operator === '=');
    if (raw === '' || raw === null) {
      if (value === '=') return input === null;
      if (value === '<>') return input !== null;
      return operator === '=' && (input === null || input === '');
    }
    if (boolean !== undefined || typeof raw === 'boolean') {
      return typeof input === 'boolean' ? matchesComparison(compare(input, boolean ?? raw), operator) : operator === '<>';
    }
    if (Number.isFinite(numeric)) {
      const translated = input === null || typeof input === 'boolean' ? NaN : number(input);
      return Number.isFinite(translated) ? matchesComparison(translated - numeric, operator) : operator === '<>';
    }
    return typeof input === 'string' ? matchesComparison(compare(input, raw), operator) : operator === '<>';
  };
}

export function sumIfRange(criteria: Range, sum: Range): Range | undefined {
  const bottom = sum.top + criteria.bottom - criteria.top;
  const right = sum.left + criteria.right - criteria.left;
  if (bottom > documentLimits.rows || right > documentLimits.columns) return undefined;
  return { ...sum, bottom, right };
}

function conditional(name: 'COUNTIF' | 'SUMIF', args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
  if (args.length < 2 || args.length > (name === 'COUNTIF' ? 2 : 3)) return error('#VALUE!');
  const input = ctx.evaluate(args[0]!);
  const criterion = ctx.scalar(ctx.evaluate(args[1]!));
  const target = args[2] ? ctx.evaluate(args[2]) : input;
  if (isError(criterion)) return criterion;
  if (!isRange(input)) return !isArray(input) && isError(input) ? input : error('#VALUE!');
  if (!isRange(target)) return !isArray(target) && isError(target) ? target : error('#VALUE!');
  if (input.length !== 1 || target.length !== 1) return error('#VALUE!');
  const from = input[0]!, sum = target[0]!;
  const match = criteria(criterion, ctx.number);
  if (name === 'COUNTIF') {
    const size = (from.bottom - from.top + 1) * (from.right - from.left + 1);
    let count = match(null) ? size : 0;
    const blankMatches = match(null);
    for (const value of ctx.values(from)) {
      if (value === null) continue;
      if (match(value)) { if (!blankMatches) count++; }
      else if (blankMatches) count--;
    }
    return count;
  }
  const expanded = sumIfRange(from, sum);
  if (!expanded) return error('#REF!');
  let result = 0;
  for (const { row, column, value } of ctx.entries(expanded)) {
    if (!match(ctx.read(from, from.top + row - sum.top, from.left + column - sum.left))) continue;
    if (isError(value)) return value;
    if (typeof value === 'number') result += value;
  }
  return finite(result);
}

function lookup(args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
  if (args.length < 3 || args.length > 4) return error('#VALUE!');
  const key = ctx.scalar(ctx.evaluate(args[0]!));
  const table = ctx.evaluate(args[1]!, true);
  const index = ctx.scalar(ctx.evaluate(args[2]!));
  const mode = args[3] ? ctx.scalar(ctx.evaluate(args[3])) : true;
  if (isError(key)) return key;
  if (isError(index)) return index;
  if (isError(mode)) return mode;
  if (!isRange(table) && !isArray(table)) return isError(table) ? table : error('#VALUE!');
  if (isRange(table) && table.length !== 1) return error('#VALUE!');
  const range = isRange(table) ? table[0]! : undefined;
  const rows = isArray(table) ? table.rows : range!.bottom - range!.top + 1;
  const columns = isArray(table) ? table.columns : range!.right - range!.left + 1;
  const at = (row: number, column: number) => isArray(table)
    ? table.at(row, column)
    : ctx.read(range!, range!.top + row, range!.left + column);
  const offset = Math.trunc(ctx.number(index)) - 1;
  const approximate = logicalValue(mode);
  if (Number.isNaN(offset) || offset < 0 || approximate === undefined) return error('#VALUE!');
  if (offset >= columns) return error('#REF!');
  let found: number | undefined;
  if (key === null) return error('#N/A');
  for (let row = 0; row < rows; row++) {
    const candidate = at(row, 0);
    if (candidate === null || isError(candidate)) continue;
    if (approximate && typeof key !== typeof candidate) continue;
    const comparison = compare(candidate, key);
    if (!approximate && comparison === 0) {
      found = row;
      break;
    }
    if (approximate && comparison <= 0) found = row;
    if (approximate && comparison > 0) break;
  }
  return found === undefined ? error('#N/A') : at(found, offset);
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
    const value = ctx.evaluate(arg);
    if (!isRange(value) || value.length !== 1) return error('#VALUE!');
    ranges.push(value[0]!);
  }
  function* eligible() {
    for (const range of ranges) yield* ctx.values(range, { skipHidden: code > 100 });
  }
  return aggregate(operation, eligible());
}

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

export function evaluateFunction(name: string, args: readonly FormulaExpression[], ctx: FunctionContext): FormulaValue {
  if (!supportedFunctions.has(name)) return error('#NAME?');
  const scalar = (index: number) => ctx.scalar(ctx.evaluate(args[index]!));
  if (Object.hasOwn(aggregateCodes, name)) {
    if (!args.length) return error('#VALUE!');
    const code = aggregateCodes[name as keyof typeof aggregateCodes];
    return aggregate(code, aggregateInputs(args, ctx, code));
  }
  switch (name as NativeFunction) {
    case 'IF': {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const test = scalar(0);
      if (isError(test)) return test;
      const condition = logicalValue(test);
      if (condition === undefined) return error('#VALUE!');
      const chosen = args[condition ? 1 : 2];
      return chosen ? ctx.evaluate(chosen) : false;
    }
    case 'IFERROR': {
      if (args.length !== 2) return error('#VALUE!');
      const first = scalar(0);
      return isError(first) ? scalar(1) : first;
    }
    case 'AND':
    case 'OR': {
      let found = false, result = name === 'AND';
      for (const arg of args) {
        const input = ctx.evaluate(arg, true);
        let values: Iterable<Scalar>;
        if (isRange(input)) {
          values = (function* () {
            for (const range of input) yield* ctx.values(range);
          })();
        } else if (isArray(input)) values = arrayValues(input);
        else values = [input];
        for (const value of values) {
          if (isArray(input) && (value === null || typeof value === 'string' || isError(value))) continue;
          if (isError(value)) return value;
          if (isRange(input) && (value === null || typeof value === 'string')) continue;
          const logical = logicalValue(value);
          if (logical === undefined) return error('#VALUE!');
          found = true;
          result = name === 'AND' ? result && logical : result || logical;
        }
      }
      return found ? result : error('#VALUE!');
    }
    case 'SUMIF':
    case 'COUNTIF': return conditional(name as 'SUMIF' | 'COUNTIF', args, ctx);
    case 'VLOOKUP': return lookup(args, ctx);
    case 'SUBTOTAL': return subtotal(args, ctx);
  }
  const values: Scalar[] = [];
  const stringFunction = ['LEN', 'TRIM', 'UPPER', 'LOWER', 'LEFT', 'RIGHT', 'MID'].includes(name);
  for (let i = 0; i < args.length; i++) {
    const input = ctx.evaluate(args[i]!);
    if (stringFunction && isArray(input) && input.rows * input.columns > 1) return error('#VALUE!');
    const value = ctx.scalar(input);
    if (isError(value)) return value;
    values.push(value);
  }
  const first = values[0] ?? null;
  if (['DATE', 'YEAR', 'MONTH', 'DAY', 'TODAY'].includes(name)) return dateFunction(name, values, ctx.today, ctx.number, ctx.locale);
  if (name === 'NOT' && values.length === 1) {
    const bool = logicalValue(first);
    return bool === undefined ? error('#VALUE!') : !bool;
  }
  if (name === 'ABS' && values.length === 1) {
    const number = ctx.number(first);
    return Number.isNaN(number) ? error('#VALUE!') : finite(Math.abs(number));
  }
  if (name === 'ROUND' || name === 'ROUNDUP' || name === 'ROUNDDOWN') {
    if (values.length !== 2) return error('#VALUE!');
    const number = ctx.number(first), places = ctx.number(values[1]!);
    return Number.isNaN(number) || Number.isNaN(places) ? error('#VALUE!') :
      Number.isFinite(number) && Number.isFinite(places) ? round(name, number, places) : error('#NUM!');
  }
  if (isError(first)) return first;
  const text = ctx.text(first);
  switch (name) {
    case 'LEN': return values.length === 1 ? text.length : error('#VALUE!');
    case 'TRIM': return values.length === 1 ? text.replace(/^ +| +$/gu, '').replace(/ +/gu, ' ') : error('#VALUE!');
    case 'UPPER': return values.length === 1 ? text.toUpperCase() : error('#VALUE!');
    case 'LOWER': return values.length === 1 ? text.toLowerCase() : error('#VALUE!');
    case 'LEFT':
    case 'RIGHT': {
      if (values.length < 1 || values.length > 2) return error('#VALUE!');
      const size = values.length === 1 ? 1 : Math.trunc(ctx.number(values[1]!));
      if (!Number.isFinite(size) || size < 0) return error('#VALUE!');
      return name === 'LEFT' ? text.slice(0, size) : text.slice(Math.max(0, text.length - size));
    }
    case 'MID': {
      if (values.length !== 3) return error('#VALUE!');
      const start = Math.trunc(ctx.number(values[1]!)), size = Math.trunc(ctx.number(values[2]!));
      return start >= 1 && size >= 0 && Number.isFinite(start + size) ? text.slice(start - 1, start - 1 + size) : error('#VALUE!');
    }
    default: return error('#VALUE!');
  }
}
