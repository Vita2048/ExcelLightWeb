import type { FormulaExpression } from '../../core/formula-references';
import { error, isArray, isError, isRange, type FormulaValue, type Scalar } from '../formula-values';
import type { FunctionContext, FunctionSpec } from './types';

const errorTypeCodes: Readonly<Record<string, number>> = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
  '#GETTING_DATA': 8,
  '#SPILL!': 9,
  '#CONNECT!': 10,
  '#BLOCKED!': 11,
  '#UNKNOWN!': 12,
  '#FIELD!': 13,
  '#CALC!': 14,
};

function mapInfoPredicate(
  args: readonly FormulaExpression[],
  ctx: FunctionContext,
  predicate: (value: Scalar) => boolean,
): FormulaValue {
  if (args.length !== 1) return error('#VALUE!');
  const input = ctx.evaluate(args[0]!, true);
  if (isRange(input)) {
    if (input.length !== 1) return predicate(error('#VALUE!'));
    const area = input[0]!;
    const rows = area.bottom - area.top + 1;
    const columns = area.right - area.left + 1;
    if (rows === 1 && columns === 1) return predicate(ctx.read(area, area.top, area.left));
    return { kind: 'array', rows, columns, at: (r, c) => predicate(ctx.read(area, area.top + r, area.left + c)) };
  }
  if (isArray(input)) {
    if (input.rows === 1 && input.columns === 1) return predicate(input.at(0, 0));
    return { kind: 'array', rows: input.rows, columns: input.columns, at: (r, c) => predicate(input.at(r, c)) };
  }
  return predicate(ctx.scalar(input));
}

function parityTest(args: readonly FormulaExpression[], ctx: FunctionContext, wantEven: boolean): Scalar {
  if (args.length !== 1) return error('#VALUE!');
  const value = ctx.scalar(ctx.evaluate(args[0]!));
  if (isError(value)) return value;
  if (typeof value === 'boolean') return error('#VALUE!');
  const numeric = value === null ? 0 : ctx.number(value);
  if (!Number.isFinite(numeric)) return error('#VALUE!');
  const truncated = Math.trunc(numeric);
  return wantEven ? truncated % 2 === 0 : Math.abs(truncated % 2) === 1;
}

export const infoFunctions: readonly FunctionSpec[] = [
  {
    name: 'ISBLANK',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate: (args, ctx) => mapInfoPredicate(args, ctx, value => value === null),
  },
  {
    name: 'ISERR',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate: (args, ctx) => mapInfoPredicate(args, ctx, value => isError(value) && value.value !== '#N/A'),
  },
  {
    name: 'ISERROR',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate: (args, ctx) => mapInfoPredicate(args, ctx, value => isError(value)),
  },
  {
    name: 'ISLOGICAL',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate: (args, ctx) => mapInfoPredicate(args, ctx, value => typeof value === 'boolean'),
  },
  {
    name: 'ISNA',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate: (args, ctx) => mapInfoPredicate(args, ctx, value => isError(value) && value.value === '#N/A'),
  },
  {
    name: 'ISNONTEXT',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate: (args, ctx) => mapInfoPredicate(args, ctx, value => typeof value !== 'string'),
  },
  {
    name: 'ISNUMBER',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate: (args, ctx) => mapInfoPredicate(args, ctx, value => typeof value === 'number'),
  },
  {
    name: 'ISTEXT',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate: (args, ctx) => mapInfoPredicate(args, ctx, value => typeof value === 'string'),
  },
  {
    name: 'ISFORMULA',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['reference']],
    evaluate(args, ctx): FormulaValue {
      if (args.length !== 1) return error('#VALUE!');
      const input = ctx.evaluate(args[0]!, true);
      if (isError(input)) return input;
      if (!isRange(input) || input.length !== 1) return error('#VALUE!');
      const area = input[0]!;
      const rows = area.bottom - area.top + 1;
      const columns = area.right - area.left + 1;
      const check = (r: number, c: number) => ctx.cellAt(area.sheetId, r, c)?.content.kind === 'formula';
      if (rows === 1 && columns === 1) return check(area.top, area.left);
      return { kind: 'array', rows, columns, at: (r, c) => check(area.top + r, area.left + c) };
    },
  },
  {
    name: 'ISREF',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 1) return error('#VALUE!');
      return isRange(ctx.evaluate(args[0]!, true));
    },
  },
  {
    name: 'ISEVEN',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['number']],
    evaluate: (args, ctx) => parityTest(args, ctx, true),
  },
  {
    name: 'ISODD',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['number']],
    evaluate: (args, ctx) => parityTest(args, ctx, false),
  },
  {
    name: 'ERROR.TYPE',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['error_val']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 1) return error('#VALUE!');
      const value = ctx.scalar(ctx.evaluate(args[0]!));
      if (!isError(value)) return error('#N/A');
      return errorTypeCodes[value.value] ?? error('#N/A');
    },
  },
  {
    name: 'N',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 1) return error('#VALUE!');
      const input = ctx.evaluate(args[0]!, true);
      let scalar: Scalar;
      if (isRange(input)) {
        if (input.length !== 1) return error('#VALUE!');
        const area = input[0]!;
        scalar = ctx.read(area, area.top, area.left);
      } else if (isArray(input)) {
        scalar = input.at(0, 0);
      } else {
        scalar = ctx.scalar(input);
      }
      if (isError(scalar)) return scalar;
      if (typeof scalar === 'number') return scalar;
      if (typeof scalar === 'boolean') return scalar ? 1 : 0;
      return 0;
    },
  },
  {
    name: 'NA',
    category: 'info',
    minArgs: 0,
    maxArgs: 0,
    signature: [[]],
    evaluate: args => args.length === 0 ? error('#N/A') : error('#VALUE!'),
  },
  {
    name: 'TYPE',
    category: 'info',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 1) return error('#VALUE!');
      const input = ctx.evaluate(args[0]!, true);
      if (isArray(input)) return 64;
      if (isRange(input)) {
        if (input.length > 1) return 64;
        const area = input[0]!;
        if (area.bottom > area.top || area.right > area.left) return 64;
        const cell = ctx.read(area, area.top, area.left);
        if (isError(cell)) return 16;
        if (typeof cell === 'string') return 2;
        if (typeof cell === 'boolean') return 4;
        return 1;
      }
      const scalar = ctx.scalar(input);
      if (isError(scalar)) return 16;
      if (typeof scalar === 'string') return 2;
      if (typeof scalar === 'boolean') return 4;
      return 1;
    },
  },
];
