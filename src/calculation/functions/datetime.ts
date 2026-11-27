import { dateFunction } from '../formula-dates';
import { evaluateScalarArgs } from './helpers';
import type { FunctionSignature, FunctionSpec } from './types';

const datePart: FunctionSignature = [['serial_number']];

function dateSpec(name: 'DATE' | 'YEAR' | 'MONTH' | 'DAY' | 'TODAY', minArgs: number, maxArgs: number,
  signature: FunctionSignature, volatile = false): FunctionSpec {
  return {
    name,
    category: 'datetime',
    minArgs,
    maxArgs,
    ...(volatile && { volatile: true }),
    signature,
    evaluate(args, ctx) {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      return dateFunction(name, values, ctx.today, ctx.number, ctx.locale);
    },
  };
}

export const datetimeFunctions: readonly FunctionSpec[] = [
  dateSpec('DATE', 3, 3, [['year', 'month', 'day']]),
  dateSpec('YEAR', 1, 1, datePart),
  dateSpec('MONTH', 1, 1, datePart),
  dateSpec('DAY', 1, 1, datePart),
  dateSpec('TODAY', 0, 0, [[]], true),
];
