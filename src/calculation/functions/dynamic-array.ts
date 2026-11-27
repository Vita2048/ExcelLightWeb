import { documentLimits } from '../../core/workbook';
import {
  arrayFromMatrix, error, isArray, isError, isLambda, isOmitted, isRange,
  type FormulaArray, type FormulaError, type FormulaValue, type Scalar,
} from '../formula-values';
import type { FunctionContext, FunctionSpec } from './types';

function toArrayOrError(value: FormulaValue, ctx: FunctionContext): FormulaArray | FormulaError {
  if (isOmitted(value) || isLambda(value)) return error('#VALUE!');
  if (isArray(value)) return value;
  if (isRange(value)) {
    const converted = ctx.arrayOf(value);
    return converted ?? error('#VALUE!');
  }
  if (isError(value)) return value;
  return arrayFromMatrix(1, 1, [value]);
}

function scalarResult(value: FormulaValue, ctx: FunctionContext): Scalar {
  if (isLambda(value) || isOmitted(value)) return error('#CALC!');
  return ctx.scalar(value);
}

export const dynamicArrayFunctions: readonly FunctionSpec[] = [
  {
    name: 'LET',
    category: 'dynamic-array',
    minArgs: 3,
    maxArgs: 253,
    signature: [['name1', 'name_value1', 'calculation_or_name2'], true],
    evaluate(args, ctx) {
      if (args.length < 3 || args.length % 2 === 0) return error('#VALUE!');
      const localEnv = new Map(ctx.env);
      for (let i = 0; i < args.length - 1; i += 2) {
        const nameNode = args[i]!;
        if (nameNode.kind !== 'name') return error('#VALUE!');
        const val = ctx.evaluateInEnv(args[i + 1]!, localEnv, true);
        if (isOmitted(val)) return error('#VALUE!');
        localEnv.set(nameNode.name.toUpperCase(), val);
      }
      return ctx.evaluateInEnv(args.at(-1)!, localEnv, true);
    },
  },
  {
    name: 'LAMBDA',
    category: 'dynamic-array',
    minArgs: 1,
    maxArgs: 253,
    signature: [['parameter_or_calculation'], true],
    evaluate(args, ctx) {
      if (args.length < 1) return error('#VALUE!');
      const params: string[] = [];
      const seen = new Set<string>();
      for (let i = 0; i < args.length - 1; i++) {
        const param = args[i]!;
        if (param.kind !== 'name') return error('#VALUE!');
        const key = param.name.toUpperCase();
        if (seen.has(key)) return error('#VALUE!');
        seen.add(key);
        params.push(key);
      }
      return {
        kind: 'lambda',
        params,
        body: args.at(-1)!,
        env: ctx.env,
      };
    },
  },
  {
    name: 'ISOMITTED',
    category: 'dynamic-array',
    minArgs: 1,
    maxArgs: 1,
    signature: [['argument']],
    evaluate(args, ctx) {
      if (args.length !== 1) return error('#VALUE!');
      return isOmitted(ctx.evaluate(args[0]!, true));
    },
  },
  {
    name: 'MAP',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 255,
    signature: [['array1', 'lambda_or_array2'], true],
    evaluate(args, ctx) {
      if (args.length < 2) return error('#VALUE!');
      const fnVal = ctx.evaluate(args.at(-1)!, true);
      if (!isLambda(fnVal)) return isError(fnVal) ? fnVal : error('#VALUE!');
      const arrays: FormulaArray[] = [];
      for (let i = 0; i < args.length - 1; i++) {
        const arr = toArrayOrError(ctx.evaluate(args[i]!, true), ctx);
        if ('kind' in arr && arr.kind === 'error') return arr;
        arrays.push(arr as FormulaArray);
      }
      if (fnVal.params.length !== arrays.length) return error('#VALUE!');
      const rows = Math.max(...arrays.map(a => a.rows));
      const columns = Math.max(...arrays.map(a => a.columns));
      if (rows * columns > 100_000) return error('#NUM!');
      const cells: Scalar[] = [];
      for (let r = 0; r < rows; r++) for (let c = 0; c < columns; c++) {
        const callArgs: FormulaValue[] = arrays.map(arr =>
          r < arr.rows && c < arr.columns ? arr.at(r, c) : error('#N/A'));
        cells.push(scalarResult(ctx.invokeLambda(fnVal, callArgs), ctx));
      }
      return arrayFromMatrix(rows, columns, cells);
    },
  },
  {
    name: 'REDUCE',
    category: 'dynamic-array',
    minArgs: 3,
    maxArgs: 3,
    signature: [['[initial_value]', 'array', 'lambda']],
    evaluate(args, ctx) {
      if (args.length !== 3) return error('#VALUE!');
      const initVal = ctx.evaluate(args[0]!, true);
      const arr = toArrayOrError(ctx.evaluate(args[1]!, true), ctx);
      if ('kind' in arr && arr.kind === 'error') return arr;
      const fnVal = ctx.evaluate(args[2]!, true);
      if (!isLambda(fnVal)) return isError(fnVal) ? fnVal : error('#VALUE!');
      if (fnVal.params.length !== 2) return error('#VALUE!');

      let acc: FormulaValue = initVal;
      const useFirstAsInit = isOmitted(initVal);
      for (let r = 0; r < (arr as FormulaArray).rows; r++) {
        for (let c = 0; c < (arr as FormulaArray).columns; c++) {
          const item = (arr as FormulaArray).at(r, c);
          if (useFirstAsInit && r === 0 && c === 0) acc = item;
          else acc = ctx.invokeLambda(fnVal, [acc, item]);
        }
      }
      return isLambda(acc) || isOmitted(acc) ? error('#CALC!') : acc;
    },
  },
  {
    name: 'SCAN',
    category: 'dynamic-array',
    minArgs: 3,
    maxArgs: 3,
    signature: [['[initial_value]', 'array', 'lambda']],
    evaluate(args, ctx) {
      if (args.length !== 3) return error('#VALUE!');
      const initVal = ctx.evaluate(args[0]!, true);
      const arr = toArrayOrError(ctx.evaluate(args[1]!, true), ctx);
      if ('kind' in arr && arr.kind === 'error') return arr;
      const fnVal = ctx.evaluate(args[2]!, true);
      if (!isLambda(fnVal)) return isError(fnVal) ? fnVal : error('#VALUE!');
      if (fnVal.params.length !== 2) return error('#VALUE!');

      const source = arr as FormulaArray;
      let acc: FormulaValue = initVal;
      const useFirstAsInit = isOmitted(initVal);
      const out: Scalar[] = [];
      for (let r = 0; r < source.rows; r++) for (let c = 0; c < source.columns; c++) {
        const item = source.at(r, c);
        if (useFirstAsInit && r === 0 && c === 0) acc = item;
        else acc = ctx.invokeLambda(fnVal, [acc, item]);
        out.push(scalarResult(acc, ctx));
      }
      return arrayFromMatrix(source.rows, source.columns, out);
    },
  },
  {
    name: 'BYROW',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array', 'lambda']],
    evaluate(args, ctx) {
      if (args.length !== 2) return error('#VALUE!');
      const arr = toArrayOrError(ctx.evaluate(args[0]!, true), ctx);
      if ('kind' in arr && arr.kind === 'error') return arr;
      const fnVal = ctx.evaluate(args[1]!, true);
      if (!isLambda(fnVal)) return isError(fnVal) ? fnVal : error('#VALUE!');
      if (fnVal.params.length !== 1) return error('#VALUE!');

      const source = arr as FormulaArray;
      const out: Scalar[] = [];
      for (let r = 0; r < source.rows; r++) {
        const rowArr: FormulaArray = {
          kind: 'array',
          rows: 1,
          columns: source.columns,
          at: (_, col) => source.at(r, col),
        };
        const res = ctx.invokeLambda(fnVal, [rowArr]);
        if (isArray(res) && res.rows * res.columns > 1) return error('#CALC!');
        if (isRange(res) && (res.length !== 1 || res[0]!.bottom !== res[0]!.top || res[0]!.right !== res[0]!.left)) {
          return error('#CALC!');
        }
        out.push(scalarResult(res, ctx));
      }
      return arrayFromMatrix(source.rows, 1, out);
    },
  },
  {
    name: 'BYCOL',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 2,
    signature: [['array', 'lambda']],
    evaluate(args, ctx) {
      if (args.length !== 2) return error('#VALUE!');
      const arr = toArrayOrError(ctx.evaluate(args[0]!, true), ctx);
      if ('kind' in arr && arr.kind === 'error') return arr;
      const fnVal = ctx.evaluate(args[1]!, true);
      if (!isLambda(fnVal)) return isError(fnVal) ? fnVal : error('#VALUE!');
      if (fnVal.params.length !== 1) return error('#VALUE!');

      const source = arr as FormulaArray;
      const out: Scalar[] = [];
      for (let c = 0; c < source.columns; c++) {
        const colArr: FormulaArray = {
          kind: 'array',
          rows: source.rows,
          columns: 1,
          at: row => source.at(row, c),
        };
        const res = ctx.invokeLambda(fnVal, [colArr]);
        if (isArray(res) && res.rows * res.columns > 1) return error('#CALC!');
        if (isRange(res) && (res.length !== 1 || res[0]!.bottom !== res[0]!.top || res[0]!.right !== res[0]!.left)) {
          return error('#CALC!');
        }
        out.push(scalarResult(res, ctx));
      }
      return arrayFromMatrix(1, source.columns, out);
    },
  },
  {
    name: 'MAKEARRAY',
    category: 'dynamic-array',
    minArgs: 3,
    maxArgs: 3,
    signature: [['rows', 'cols', 'lambda']],
    evaluate(args, ctx) {
      if (args.length !== 3) return error('#VALUE!');
      const rVal = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(rVal)) return rVal;
      const cVal = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(cVal)) return cVal;
      const fnVal = ctx.evaluate(args[2]!, true);
      if (!isLambda(fnVal)) return isError(fnVal) ? fnVal : error('#VALUE!');
      if (fnVal.params.length !== 2) return error('#VALUE!');

      const rows = Math.trunc(ctx.number(rVal));
      const cols = Math.trunc(ctx.number(cVal));
      if (!Number.isFinite(rows) || !Number.isFinite(cols) || rows < 1 || cols < 1) return error('#VALUE!');
      if (rows > documentLimits.rows || cols > documentLimits.columns || rows * cols > 100_000) return error('#NUM!');

      const out: Scalar[] = [];
      for (let r = 1; r <= rows; r++) for (let c = 1; c <= cols; c++) {
        const res = ctx.invokeLambda(fnVal, [r, c]);
        if (isArray(res) && res.rows * res.columns > 1) return error('#CALC!');
        out.push(scalarResult(res, ctx));
      }
      return arrayFromMatrix(rows, cols, out);
    },
  },
];
