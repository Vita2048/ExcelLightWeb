import type { FormulaExpression } from '../../core/formula-references';
import { arrayValues, error, isArray, isError, isRange, type Scalar } from '../formula-values';
import { evaluateScalarArgs, logicalValue } from './helpers';
import type { FunctionContext, FunctionSignature, FunctionSpec } from './types';

const logicals: FunctionSignature = [['logical1', '[logical2]'], true];

function evaluateAndOr(kind: 'AND' | 'OR', args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
  let found = false, result = kind === 'AND';
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
      result = kind === 'AND' ? result && logical : result || logical;
    }
  }
  return found ? result : error('#VALUE!');
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
      return chosen ? ctx.evaluate(chosen) : false;
    },
  },
  {
    name: 'AND',
    category: 'logical',
    minArgs: 1,
    maxArgs: 255,
    signature: logicals,
    evaluate: (args, ctx) => evaluateAndOr('AND', args, ctx),
  },
  {
    name: 'OR',
    category: 'logical',
    minArgs: 1,
    maxArgs: 255,
    signature: logicals,
    evaluate: (args, ctx) => evaluateAndOr('OR', args, ctx),
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
];
