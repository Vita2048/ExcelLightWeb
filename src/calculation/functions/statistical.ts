import { error } from '../formula-values';
import { aggregate, aggregateInputs, conditional } from './helpers';
import type { FunctionSignature, FunctionSpec } from './types';

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
];
