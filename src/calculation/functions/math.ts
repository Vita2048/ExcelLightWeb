import type { FormulaExpression } from '../../core/formula-references';
import { error, finite, isError, isRange, type Range, type Scalar } from '../formula-values';
import { aggregate, aggregateInputs, conditional, evaluateScalarArgs } from './helpers';
import type { FunctionContext, FunctionSignature, FunctionSpec } from './types';

const rounding: FunctionSignature = [['number', 'num_digits']];

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
    name: 'ABS',
    category: 'math',
    minArgs: 1,
    maxArgs: 1,
    signature: [['number']],
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
    name: 'SUMIF',
    category: 'math',
    minArgs: 2,
    maxArgs: 3,
    signature: [['range', 'criteria', '[sum_range]']],
    evaluate: (args, ctx) => conditional('SUMIF', args, ctx),
  },
  {
    name: 'SUBTOTAL',
    category: 'math',
    minArgs: 2,
    maxArgs: 255,
    signature: [['function_num', 'ref1', '[ref2]'], true],
    evaluate: subtotal,
  },
];
