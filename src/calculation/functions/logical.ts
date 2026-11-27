import type { FormulaExpression } from '../../core/formula-references';
import {
  arrayFromMatrix, arrayValues, compare, error, isArray, isError, isRange,
  type FormulaValue, type Scalar,
} from '../formula-values';
import { evaluateScalarArgs, logicalValue } from './helpers';
import type { FunctionContext, FunctionSignature, FunctionSpec } from './types';

const logicals: FunctionSignature = [['logical1', '[logical2]'], true];

function evaluateLogicalFold(kind: 'AND' | 'OR' | 'XOR', args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
  let found = false;
  let result = kind === 'AND';
  let trueCount = 0;
  for (const arg of args) {
    const input = ctx.evaluate(arg, true);
    let values: Iterable<Scalar>;
    if (isRange(input)) {
      values = (function* () {
        for (const range of input) yield* ctx.values(range);
      })();
    } else if (isArray(input)) values = arrayValues(input);
    else values = [ctx.scalar(input)];
    for (const value of values) {
      if (isArray(input) && (value === null || typeof value === 'string' || isError(value))) continue;
      if (isError(value)) return value;
      if ((isRange(input) || kind === 'XOR') && (value === null || typeof value === 'string')) continue;
      const logical = logicalValue(value);
      if (logical === undefined) return error('#VALUE!');
      found = true;
      if (kind === 'AND') result = result && logical;
      else if (kind === 'OR') result = result || logical;
      else if (logical) trueCount++;
    }
  }
  if (!found) return error('#VALUE!');
  return kind === 'XOR' ? trueCount % 2 === 1 : result;
}

export const logicalFunctions: readonly FunctionSpec[] = [
  {
    name: 'IF',
    category: 'logical',
    minArgs: 2,
    maxArgs: 3,
    signature: [['logical_test', '[value_if_true]', '[value_if_false]']],
    evaluate(args, ctx) {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const test = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(test)) return test;
      const condition = logicalValue(test);
      if (condition === undefined) return error('#VALUE!');
      const chosen = args[condition ? 1 : 2];
      return chosen ? ctx.evaluate(chosen, true) : false;
    },
  },
  {
    name: 'AND',
    category: 'logical',
    minArgs: 1,
    maxArgs: 255,
    signature: logicals,
    evaluate: (args, ctx) => evaluateLogicalFold('AND', args, ctx),
  },
  {
    name: 'OR',
    category: 'logical',
    minArgs: 1,
    maxArgs: 255,
    signature: logicals,
    evaluate: (args, ctx) => evaluateLogicalFold('OR', args, ctx),
  },
  {
    name: 'XOR',
    category: 'logical',
    minArgs: 1,
    maxArgs: 255,
    signature: logicals,
    evaluate: (args, ctx) => evaluateLogicalFold('XOR', args, ctx),
  },
  {
    name: 'NOT',
    category: 'logical',
    minArgs: 1,
    maxArgs: 1,
    signature: [['logical']],
    evaluate(args, ctx) {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const bool = logicalValue(values[0] ?? null);
      return bool === undefined ? error('#VALUE!') : !bool;
    },
  },
  {
    name: 'IFERROR',
    category: 'logical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['value', 'value_if_error']],
    evaluate(args, ctx) {
      if (args.length !== 2) return error('#VALUE!');
      const first = ctx.scalar(ctx.evaluate(args[0]!));
      return isError(first) ? ctx.scalar(ctx.evaluate(args[1]!)) : first;
    },
  },
  {
    name: 'IFNA',
    category: 'logical',
    minArgs: 2,
    maxArgs: 2,
    signature: [['value', 'value_if_na']],
    evaluate(args, ctx): FormulaValue {
      if (args.length !== 2) return error('#VALUE!');
      const value = ctx.evaluate(args[0]!, true);
      const arr = ctx.arrayOf(value);
      if (arr && arr.rows * arr.columns > 1) {
        let hasNa = false;
        const cells: Scalar[] = new Array(arr.rows * arr.columns);
        for (let r = 0; r < arr.rows; r++) for (let c = 0; c < arr.columns; c++) {
          const item = arr.at(r, c);
          if (isError(item) && item.value === '#N/A') hasNa = true;
          cells[r * arr.columns + c] = item;
        }
        if (!hasNa) return arr;
        const fallback = ctx.scalar(ctx.evaluate(args[1]!));
        for (let i = 0; i < cells.length; i++) {
          const item = cells[i]!;
          if (isError(item) && item.value === '#N/A') cells[i] = fallback;
        }
        return arrayFromMatrix(arr.rows, arr.columns, cells);
      }
      const scalarVal = ctx.scalar(value);
      return isError(scalarVal) && scalarVal.value === '#N/A' ? ctx.evaluate(args[1]!, true) : value;
    },
  },
  {
    name: 'IFS',
    category: 'logical',
    minArgs: 2,
    maxArgs: 254,
    signature: [['logical_test1', 'value_if_true1', '[logical_test2]', '[value_if_true2]'], true],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2 || args.length % 2 !== 0) return error('#VALUE!');
      for (let i = 0; i < args.length; i += 2) {
        const test = ctx.scalar(ctx.evaluate(args[i]!));
        if (isError(test)) return test;
        const condition = logicalValue(test);
        if (condition === undefined) return error('#VALUE!');
        if (condition) return ctx.evaluate(args[i + 1]!, true);
      }
      return error('#N/A');
    },
  },
  {
    name: 'SWITCH',
    category: 'logical',
    minArgs: 3,
    maxArgs: 254,
    signature: [['expression', 'value1', 'result1', '[default_or_value2]'], true],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 3) return error('#VALUE!');
      const target = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(target)) return target;
      let i = 1;
      for (; i + 1 < args.length; i += 2) {
        const candidate = ctx.scalar(ctx.evaluate(args[i]!));
        if (isError(candidate)) return candidate;
        if (compare(target, candidate) === 0) return ctx.evaluate(args[i + 1]!, true);
      }
      return i < args.length ? ctx.evaluate(args[i]!, true) : error('#N/A');
    },
  },
  {
    name: 'TRUE',
    category: 'logical',
    minArgs: 0,
    maxArgs: 0,
    signature: [[]],
    evaluate: args => args.length === 0 ? true : error('#VALUE!'),
  },
  {
    name: 'FALSE',
    category: 'logical',
    minArgs: 0,
    maxArgs: 0,
    signature: [[]],
    evaluate: args => args.length === 0 ? false : error('#VALUE!'),
  },
];
