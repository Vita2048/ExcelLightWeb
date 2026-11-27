import type { FormulaExpression } from '../../core/formula-references';
import { documentLimits } from '../../core/workbook';
import { arrayValues, compare, error, finite, isArray, isError, isRange, matchesComparison,
  type FormulaError, type Range, type Scalar } from '../formula-values';
import type { FunctionContext } from './types';

export function logicalValue(value: Scalar): boolean | undefined {
  if (typeof value === 'string') {
    const key = value.trim().toUpperCase();
    return key === 'TRUE' ? true : key === 'FALSE' ? false : undefined;
  }
  return isError(value) ? undefined : Boolean(value);
}

export function aggregate(code: number, values: Iterable<Scalar>): Scalar {
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

export function* aggregateInputs(args: readonly FormulaExpression[], ctx: FunctionContext, code: number): Iterable<Scalar> {
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

const comparisons = /^(<=|>=|<>|=|<|>)/u;

export function wildcard(text: string): (value: string) => boolean {
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

export function criteria(value: Exclude<Scalar, FormulaError>, number: FunctionContext['number']): (input: Scalar) => boolean {
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

export function sumIfRange(criteriaRange: Range, sumRange: Range): Range | undefined {
  const bottom = sumRange.top + criteriaRange.bottom - criteriaRange.top;
  const right = sumRange.left + criteriaRange.right - criteriaRange.left;
  if (bottom > documentLimits.rows || right > documentLimits.columns) return undefined;
  return { ...sumRange, bottom, right };
}

export function conditional(name: 'COUNTIF' | 'SUMIF', args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
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

export function evaluateScalarArgs(
  args: readonly FormulaExpression[],
  ctx: FunctionContext,
  rejectMultiCellArray = false,
): Scalar[] | FormulaError {
  const values: Scalar[] = [];
  for (let i = 0; i < args.length; i++) {
    const input = ctx.evaluate(args[i]!);
    if (rejectMultiCellArray && isArray(input) && input.rows * input.columns > 1) return error('#VALUE!');
    const value = ctx.scalar(input);
    if (isError(value)) return value;
    values.push(value);
  }
  return values;
}
