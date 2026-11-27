import { error, isError, type Scalar } from '../formula-values';
import { evaluateScalarArgs } from './helpers';
import type { FunctionSignature, FunctionSpec } from './types';

const edgeText: FunctionSignature = [['text', '[num_chars]']];
const singleText: FunctionSignature = [['text']];

function singleTextSpec(name: 'LEN' | 'TRIM' | 'UPPER' | 'LOWER', transform: (text: string) => Scalar): FunctionSpec {
  return {
    name,
    category: 'text',
    minArgs: 1,
    maxArgs: 1,
    signature: singleText,
    evaluate(args, ctx) {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const first = values[0] ?? null;
      if (isError(first)) return first;
      return transform(ctx.text(first));
    },
  };
}

function edgeTextSpec(name: 'LEFT' | 'RIGHT'): FunctionSpec {
  return {
    name,
    category: 'text',
    minArgs: 1,
    maxArgs: 2,
    signature: edgeText,
    evaluate(args, ctx) {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length < 1 || values.length > 2) return error('#VALUE!');
      const first = values[0] ?? null;
      if (isError(first)) return first;
      const text = ctx.text(first);
      const size = values.length === 1 ? 1 : Math.trunc(ctx.number(values[1]!));
      if (!Number.isFinite(size) || size < 0) return error('#VALUE!');
      return name === 'LEFT' ? text.slice(0, size) : text.slice(Math.max(0, text.length - size));
    },
  };
}

export const textFunctions: readonly FunctionSpec[] = [
  edgeTextSpec('LEFT'),
  edgeTextSpec('RIGHT'),
  {
    name: 'MID',
    category: 'text',
    minArgs: 3,
    maxArgs: 3,
    signature: [['text', 'start_num', 'num_chars']],
    evaluate(args, ctx) {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 3) return error('#VALUE!');
      const first = values[0] ?? null;
      if (isError(first)) return first;
      const text = ctx.text(first);
      const start = Math.trunc(ctx.number(values[1]!)), size = Math.trunc(ctx.number(values[2]!));
      return start >= 1 && size >= 0 && Number.isFinite(start + size) ? text.slice(start - 1, start - 1 + size) : error('#VALUE!');
    },
  },
  singleTextSpec('LEN', text => text.length),
  singleTextSpec('TRIM', text => text.replace(/^ +| +$/gu, '').replace(/ +/gu, ' ')),
  singleTextSpec('UPPER', text => text.toUpperCase()),
  singleTextSpec('LOWER', text => text.toLowerCase()),
];
