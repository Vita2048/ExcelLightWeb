import type { FormulaExpression } from '../../core/formula-references';
import { documentLimits } from '../../core/workbook';
import {
  arrayFromMatrix, arrayValues, compare, error, finite, isError, isOmitted,
  type FormulaArray, type FormulaError, type FormulaValue, type Scalar,
} from '../formula-values';
import { logicalValue } from './helpers';
import type { FunctionContext, FunctionSpec } from './types';

function evalArray(expr: FormulaExpression, ctx: FunctionContext): FormulaArray | FormulaError {
  const val = ctx.evaluate(expr, true);
  if (isError(val)) return val;
  const arr = ctx.arrayOf(val);
  return arr ?? error('#VALUE!');
}

function intArg(expr: FormulaExpression | undefined, fallback: number, ctx: FunctionContext): number | FormulaError {
  if (!expr) return fallback;
  const val = ctx.evaluate(expr);
  if (isOmitted(val)) return fallback;
  const scalar = ctx.scalar(val);
  if (isError(scalar)) return scalar;
  const n = ctx.number(scalar);
  return Number.isFinite(n) ? Math.trunc(n) : error('#VALUE!');
}

function numArg(expr: FormulaExpression | undefined, fallback: number, ctx: FunctionContext): number | FormulaError {
  if (!expr) return fallback;
  const val = ctx.evaluate(expr);
  if (isOmitted(val)) return fallback;
  const scalar = ctx.scalar(val);
  if (isError(scalar)) return scalar;
  const n = ctx.number(scalar);
  return Number.isFinite(n) ? n : error('#VALUE!');
}

function boolArg(expr: FormulaExpression | undefined, fallback: boolean, ctx: FunctionContext): boolean | FormulaError {
  if (!expr) return fallback;
  const val = ctx.evaluate(expr);
  if (isOmitted(val)) return fallback;
  const scalar = ctx.scalar(val);
  if (isError(scalar)) return scalar;
  const bool = logicalValue(scalar);
  return bool === undefined ? error('#VALUE!') : bool;
}

function compareSortScalar(a: Scalar, b: Scalar, descending: boolean): number {
  if (a === null && b === null) return 0;
  if (a === null) return 1;
  if (b === null) return -1;
  if (isError(a) && isError(b)) return 0;
  if (isError(a)) return descending ? -1 : 1;
  if (isError(b)) return descending ? 1 : -1;
  const cmp = compare(a, b);
  return descending ? -cmp : cmp;
}

function uniqueKey(value: Scalar): string {
  if (value === null) return 'z:';
  if (isError(value)) return `e:${value.value}`;
  if (typeof value === 'number') return `n:${value === 0 ? 0 : value}`;
  if (typeof value === 'boolean') return `b:${value ? 1 : 0}`;
  return `s:${value.toUpperCase()}`;
}

function materialize(arr: FormulaArray): Scalar[] {
  const cells: Scalar[] = new Array(arr.rows * arr.columns);
  for (let r = 0; r < arr.rows; r++) for (let c = 0; c < arr.columns; c++) {
    cells[r * arr.columns + c] = arr.at(r, c);
  }
  return cells;
}

function flattenArray(
  arr: FormulaArray,
  ignore: number,
  scanByCol: boolean,
): Scalar[] | FormulaError {
  if (ignore < 0 || ignore > 3) return error('#VALUE!');
  const result: Scalar[] = [];
  const outer = scanByCol ? arr.columns : arr.rows;
  const inner = scanByCol ? arr.rows : arr.columns;
  for (let i = 0; i < outer; i++) for (let j = 0; j < inner; j++) {
    const val = scanByCol ? arr.at(j, i) : arr.at(i, j);
    if ((ignore === 1 || ignore === 3) && val === null) continue;
    if ((ignore === 2 || ignore === 3) && isError(val)) continue;
    result.push(val);
  }
  return result.length ? result : error('#CALC!');
}

function readIntList(expr: FormulaExpression, ctx: FunctionContext): number[] | FormulaError {
  const val = ctx.evaluate(expr, true);
  if (isError(val)) return val;
  const arr = ctx.arrayOf(val);
  if (!arr) return error('#VALUE!');
  const out: number[] = [];
  for (const item of arrayValues(arr)) {
    if (isError(item)) return item;
    const n = ctx.number(item);
    if (!Number.isFinite(n)) return error('#VALUE!');
    out.push(Math.trunc(n));
  }
  return out;
}

function splitMulti(text: string, delimiters: readonly string[], caseInsensitive: boolean): string[] {
  if (!delimiters.length) return [text];
  const needles = delimiters.filter(d => d.length > 0);
  if (!needles.length) return [text];
  const sorted = [...needles].sort((a, b) => b.length - a.length);
  const hay = caseInsensitive ? text.toUpperCase() : text;
  const matchedNeedles = caseInsensitive ? sorted.map(d => d.toUpperCase()) : sorted;
  const parts: string[] = [];
  let start = 0;
  for (let i = 0; i <= text.length;) {
    let matched = 0;
    for (const needle of matchedNeedles) {
      if (hay.startsWith(needle, i)) {
        matched = needle.length;
        break;
      }
    }
    if (matched > 0) {
      parts.push(text.slice(start, i));
      i += matched;
      start = i;
    } else {
      i++;
    }
  }
  parts.push(text.slice(start));
  return parts;
}

function readDelimiters(expr: FormulaExpression | undefined, ctx: FunctionContext): string[] | FormulaError {
  if (!expr) return [];
  const val = ctx.evaluate(expr, true);
  if (isOmitted(val)) return [];
  if (isError(val)) return val;
  const arr = ctx.arrayOf(val);
  if (!arr) return error('#VALUE!');
  const list: string[] = [];
  for (const item of arrayValues(arr)) {
    if (isError(item)) return item;
    if (item === null) continue;
    const s = ctx.text(item);
    if (s.length === 0) return error('#VALUE!');
    list.push(s);
  }
  return list;
}

export const shapingFunctions: readonly FunctionSpec[] = [
  {
    name: 'FILTER',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 3,
    signature: [['array', 'include', '[if_empty]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const source = evalArray(args[0]!, ctx);
      if (isError(source)) return source;
      const include = evalArray(args[1]!, ctx);
      if (isError(include)) return include;

      const byRow = include.rows === source.rows && include.columns === 1;
      const byCol = !byRow && include.columns === source.columns && include.rows === 1;
      if (!byRow && !byCol) return error('#VALUE!');

      const count = byRow ? source.rows : source.columns;
      const indices: number[] = [];
      for (let i = 0; i < count; i++) {
        const flag = byRow ? include.at(i, 0) : include.at(0, i);
        if (isError(flag)) return flag;
        const bool = logicalValue(flag);
        if (bool === undefined) return error('#VALUE!');
        if (bool) indices.push(i);
      }

      if (!indices.length) {
        if (args[2]) {
          const emptyVal = ctx.evaluate(args[2], true);
          if (!isOmitted(emptyVal)) return emptyVal;
        }
        return error('#CALC!');
      }

      if (byRow) {
        const cells: Scalar[] = [];
        for (const r of indices) for (let c = 0; c < source.columns; c++) cells.push(source.at(r, c));
        return arrayFromMatrix(indices.length, source.columns, cells);
      }
      const cells: Scalar[] = [];
      for (let r = 0; r < source.rows; r++) for (const c of indices) cells.push(source.at(r, c));
      return arrayFromMatrix(source.rows, indices.length, cells);
    },
  },
  {
    name: 'SORT',
    category: 'dynamic-array',
    minArgs: 1,
    maxArgs: 4,
    signature: [['array', '[sort_index]', '[sort_order]', '[by_col]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 1 || args.length > 4) return error('#VALUE!');
      const source = evalArray(args[0]!, ctx);
      if (isError(source)) return source;
      const byCol = boolArg(args[3], false, ctx);
      if (isError(byCol)) return byCol;

      const sortIndices = args[1] && !isOmitted(ctx.evaluate(args[1])) ? readIntList(args[1], ctx) : [1];
      if (isError(sortIndices)) return sortIndices;
      const sortOrders = args[2] && !isOmitted(ctx.evaluate(args[2])) ? readIntList(args[2], ctx) : [1];
      if (isError(sortOrders)) return sortOrders;

      const keyLimit = byCol ? source.rows : source.columns;
      for (const idx of sortIndices) {
        if (idx < 1 || idx > keyLimit) return error('#VALUE!');
      }
      for (const ord of sortOrders) {
        if (ord !== 1 && ord !== -1) return error('#VALUE!');
      }
      if (sortOrders.length > 1 && sortOrders.length !== sortIndices.length) return error('#VALUE!');

      const count = byCol ? source.columns : source.rows;
      const data = materialize(source);
      const cell = (r: number, c: number) => data[r * source.columns + c] ?? null;
      const perm = Array.from({ length: count }, (_, i) => i);
      perm.sort((a, b) => {
        for (let k = 0; k < sortIndices.length; k++) {
          const idx = sortIndices[k]! - 1;
          const desc = (sortOrders[sortOrders.length === 1 ? 0 : k] ?? 1) === -1;
          const va = byCol ? cell(idx, a) : cell(a, idx);
          const vb = byCol ? cell(idx, b) : cell(b, idx);
          const cmp = compareSortScalar(va, vb, desc);
          if (cmp !== 0) return cmp;
        }
        return a - b;
      });

      const out: Scalar[] = new Array(data.length);
      if (!byCol) {
        for (let r = 0; r < source.rows; r++) for (let c = 0; c < source.columns; c++) {
          out[r * source.columns + c] = cell(perm[r]!, c);
        }
      } else {
        for (let r = 0; r < source.rows; r++) for (let c = 0; c < source.columns; c++) {
          out[r * source.columns + c] = cell(r, perm[c]!);
        }
      }
      return arrayFromMatrix(source.rows, source.columns, out);
    },
  },
  {
    name: 'SORTBY',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 255,
    signature: [['array', 'by_array1', '[sort_order1]', '[by_array2]'], true],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2) return error('#VALUE!');
      const source = evalArray(args[0]!, ctx);
      if (isError(source)) return source;

      const by1 = evalArray(args[1]!, ctx);
      if (isError(by1)) return by1;
      const byRow = by1.rows === source.rows && by1.columns === 1;
      const byCol = !byRow && by1.columns === source.columns && by1.rows === 1;
      if (!byRow && !byCol) return error('#VALUE!');

      const keys: { values: Scalar[]; descending: boolean }[] = [];
      for (let i = 1; i < args.length; i += 2) {
        const byArr = i === 1 ? by1 : evalArray(args[i]!, ctx);
        if (isError(byArr)) return byArr;
        if (byRow ? (byArr.rows !== source.rows || byArr.columns !== 1) : (byArr.columns !== source.columns || byArr.rows !== 1)) {
          return error('#VALUE!');
        }
        const order = intArg(args[i + 1], 1, ctx);
        if (isError(order)) return order;
        if (order !== 1 && order !== -1) return error('#VALUE!');
        keys.push({ values: materialize(byArr), descending: order === -1 });
      }

      const count = byRow ? source.rows : source.columns;
      const data = materialize(source);
      const cell = (r: number, c: number) => data[r * source.columns + c] ?? null;
      const perm = Array.from({ length: count }, (_, idx) => idx);
      perm.sort((a, b) => {
        for (const key of keys) {
          const cmp = compareSortScalar(key.values[a] ?? null, key.values[b] ?? null, key.descending);
          if (cmp !== 0) return cmp;
        }
        return a - b;
      });

      const out: Scalar[] = new Array(data.length);
      if (byRow) {
        for (let r = 0; r < source.rows; r++) for (let c = 0; c < source.columns; c++) {
          out[r * source.columns + c] = cell(perm[r]!, c);
        }
      } else {
        for (let r = 0; r < source.rows; r++) for (let c = 0; c < source.columns; c++) {
          out[r * source.columns + c] = cell(r, perm[c]!);
        }
      }
      return arrayFromMatrix(source.rows, source.columns, out);
    },
  },
  {
    name: 'UNIQUE',
    category: 'dynamic-array',
    minArgs: 1,
    maxArgs: 3,
    signature: [['array', '[by_col]', '[exactly_once]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 1 || args.length > 3) return error('#VALUE!');
      const source = evalArray(args[0]!, ctx);
      if (isError(source)) return source;
      const byCol = boolArg(args[1], false, ctx);
      if (isError(byCol)) return byCol;
      const exactlyOnce = boolArg(args[2], false, ctx);
      if (isError(exactlyOnce)) return exactlyOnce;

      const data = materialize(source);
      const cell = (r: number, c: number) => data[r * source.columns + c] ?? null;
      const major = byCol ? source.columns : source.rows;
      const minor = byCol ? source.rows : source.columns;
      const seen = new Map<string, { first: number; count: number }>();

      for (let i = 0; i < major; i++) {
        const parts: string[] = [];
        for (let j = 0; j < minor; j++) {
          parts.push(uniqueKey(byCol ? cell(j, i) : cell(i, j)));
        }
        const signature = parts.join('\u0000');
        const entry = seen.get(signature);
        if (entry) entry.count++;
        else seen.set(signature, { first: i, count: 1 });
      }

      const kept: number[] = [];
      for (const entry of seen.values()) {
        if (!exactlyOnce || entry.count === 1) kept.push(entry.first);
      }
      if (!kept.length) return error('#CALC!');

      if (!byCol) {
        const out: Scalar[] = [];
        for (const r of kept) for (let c = 0; c < source.columns; c++) out.push(cell(r, c));
        return arrayFromMatrix(kept.length, source.columns, out);
      }
      const out: Scalar[] = [];
      for (let r = 0; r < source.rows; r++) for (const c of kept) out.push(cell(r, c));
      return arrayFromMatrix(source.rows, kept.length, out);
    },
  },
  {
    name: 'SEQUENCE',
    category: 'dynamic-array',
    minArgs: 1,
    maxArgs: 4,
    signature: [['rows', '[columns]', '[start]', '[step]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 1 || args.length > 4) return error('#VALUE!');
      const rows = intArg(args[0], 0, ctx);
      if (isError(rows)) return rows;
      const columns = intArg(args[1], 1, ctx);
      if (isError(columns)) return columns;
      const start = numArg(args[2], 1, ctx);
      if (isError(start)) return start;
      const step = numArg(args[3], 1, ctx);
      if (isError(step)) return step;
      if (rows < 0 || columns < 0) return error('#VALUE!');
      if (rows === 0 || columns === 0) return error('#CALC!');
      if (rows > documentLimits.rows || columns > documentLimits.columns || rows * columns > 100_000) return error('#NUM!');
      const cells: Scalar[] = new Array(rows * columns);
      for (let r = 0; r < rows; r++) for (let c = 0; c < columns; c++) {
        cells[r * columns + c] = finite(start + (r * columns + c) * step);
      }
      return arrayFromMatrix(rows, columns, cells);
    },
  },
  {
    name: 'RANDARRAY',
    category: 'dynamic-array',
    minArgs: 0,
    maxArgs: 5,
    volatile: true,
    signature: [['[rows]', '[columns]', '[min]', '[max]', '[whole_number]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length > 5) return error('#VALUE!');
      const rows = intArg(args[0], 1, ctx);
      if (isError(rows)) return rows;
      const columns = intArg(args[1], 1, ctx);
      if (isError(columns)) return columns;
      const min = numArg(args[2], 0, ctx);
      if (isError(min)) return min;
      const max = numArg(args[3], 1, ctx);
      if (isError(max)) return max;
      const whole = boolArg(args[4], false, ctx);
      if (isError(whole)) return whole;

      if (rows < 0 || columns < 0 || min > max) return error('#VALUE!');
      if (rows === 0 || columns === 0) return error('#CALC!');
      if (whole && (!Number.isInteger(min) || !Number.isInteger(max))) return error('#VALUE!');
      if (rows > documentLimits.rows || columns > documentLimits.columns || rows * columns > 100_000) return error('#NUM!');

      const cells: Scalar[] = new Array(rows * columns);
      for (let i = 0; i < cells.length; i++) {
        const u = Math.min(Math.max(ctx.random(), 0), 1 - Number.EPSILON);
        cells[i] = whole
          ? min + Math.floor(u * (max - min + 1))
          : finite(min + u * (max - min));
      }
      return arrayFromMatrix(rows, columns, cells);
    },
  },
  {
    name: 'VSTACK',
    category: 'dynamic-array',
    minArgs: 1,
    maxArgs: 255,
    signature: [['array1', '[array2]'], true],
    evaluate(args, ctx): FormulaValue {
      const arrays: FormulaArray[] = [];
      let totalRows = 0;
      let maxCols = 0;
      for (const arg of args) {
        const arr = evalArray(arg, ctx);
        if (isError(arr)) return arr;
        arrays.push(arr);
        totalRows += arr.rows;
        maxCols = Math.max(maxCols, arr.columns);
      }
      const cells: Scalar[] = [];
      const na = error('#N/A');
      for (const arr of arrays) {
        for (let r = 0; r < arr.rows; r++) for (let c = 0; c < maxCols; c++) {
          cells.push(c < arr.columns ? arr.at(r, c) : na);
        }
      }
      return arrayFromMatrix(totalRows, maxCols, cells);
    },
  },
  {
    name: 'HSTACK',
    category: 'dynamic-array',
    minArgs: 1,
    maxArgs: 255,
    signature: [['array1', '[array2]'], true],
    evaluate(args, ctx): FormulaValue {
      const arrays: FormulaArray[] = [];
      let maxRows = 0;
      let totalCols = 0;
      for (const arg of args) {
        const arr = evalArray(arg, ctx);
        if (isError(arr)) return arr;
        arrays.push(arr);
        maxRows = Math.max(maxRows, arr.rows);
        totalCols += arr.columns;
      }
      const cells: Scalar[] = [];
      const na = error('#N/A');
      for (let r = 0; r < maxRows; r++) {
        for (const arr of arrays) {
          for (let c = 0; c < arr.columns; c++) cells.push(r < arr.rows ? arr.at(r, c) : na);
        }
      }
      return arrayFromMatrix(maxRows, totalCols, cells);
    },
  },
  {
    name: 'TOROW',
    category: 'dynamic-array',
    minArgs: 1,
    maxArgs: 3,
    signature: [['array', '[ignore]', '[scan_by_column]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 1 || args.length > 3) return error('#VALUE!');
      const arr = evalArray(args[0]!, ctx);
      if (isError(arr)) return arr;
      const ignore = intArg(args[1], 0, ctx);
      if (isError(ignore)) return ignore;
      const scanByCol = boolArg(args[2], false, ctx);
      if (isError(scanByCol)) return scanByCol;
      const flat = flattenArray(arr, ignore, scanByCol);
      return isError(flat) ? flat : arrayFromMatrix(1, flat.length, flat);
    },
  },
  {
    name: 'TOCOL',
    category: 'dynamic-array',
    minArgs: 1,
    maxArgs: 3,
    signature: [['array', '[ignore]', '[scan_by_column]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 1 || args.length > 3) return error('#VALUE!');
      const arr = evalArray(args[0]!, ctx);
      if (isError(arr)) return arr;
      const ignore = intArg(args[1], 0, ctx);
      if (isError(ignore)) return ignore;
      const scanByCol = boolArg(args[2], false, ctx);
      if (isError(scanByCol)) return scanByCol;
      const flat = flattenArray(arr, ignore, scanByCol);
      return isError(flat) ? flat : arrayFromMatrix(flat.length, 1, flat);
    },
  },
  {
    name: 'WRAPROWS',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 3,
    signature: [['vector', 'wrap_count', '[pad_with]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const arr = evalArray(args[0]!, ctx);
      if (isError(arr)) return arr;
      if (arr.rows > 1 && arr.columns > 1) return error('#VALUE!');
      const count = intArg(args[1], 0, ctx);
      if (isError(count)) return count;
      if (count < 1) return error('#NUM!');
      let pad: Scalar = error('#N/A');
      if (args[2]) {
        const padVal = ctx.evaluate(args[2]);
        if (!isOmitted(padVal)) pad = ctx.scalar(padVal);
      }
      const items = [...arrayValues(arr)];
      const outRows = Math.ceil(items.length / count);
      const cells: Scalar[] = new Array(outRows * count);
      for (let i = 0; i < cells.length; i++) cells[i] = i < items.length ? items[i]! : pad;
      return arrayFromMatrix(outRows, count, cells);
    },
  },
  {
    name: 'WRAPCOLS',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 3,
    signature: [['vector', 'wrap_count', '[pad_with]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const arr = evalArray(args[0]!, ctx);
      if (isError(arr)) return arr;
      if (arr.rows > 1 && arr.columns > 1) return error('#VALUE!');
      const count = intArg(args[1], 0, ctx);
      if (isError(count)) return count;
      if (count < 1) return error('#NUM!');
      let pad: Scalar = error('#N/A');
      if (args[2]) {
        const padVal = ctx.evaluate(args[2]);
        if (!isOmitted(padVal)) pad = ctx.scalar(padVal);
      }
      const items = [...arrayValues(arr)];
      const outCols = Math.ceil(items.length / count);
      const cells: Scalar[] = new Array(count * outCols);
      for (let r = 0; r < count; r++) for (let c = 0; c < outCols; c++) {
        const idx = c * count + r;
        cells[r * outCols + c] = idx < items.length ? items[idx]! : pad;
      }
      return arrayFromMatrix(count, outCols, cells);
    },
  },
  {
    name: 'TAKE',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 3,
    signature: [['array', 'rows', '[columns]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const arr = evalArray(args[0]!, ctx);
      if (isError(arr)) return arr;
      const rowsVal = ctx.evaluate(args[1]!);
      const rowOmitted = isOmitted(rowsVal);
      const colsVal = args[2] ? ctx.evaluate(args[2]) : undefined;
      const colOmitted = !colsVal || isOmitted(colsVal);
      if (rowOmitted && colOmitted) return error('#VALUE!');

      let rowArg = arr.rows;
      if (!rowOmitted) {
        const s = ctx.scalar(rowsVal);
        if (isError(s)) return s;
        const n = ctx.number(s);
        if (!Number.isFinite(n)) return error('#VALUE!');
        rowArg = Math.trunc(n);
      }
      let colArg = arr.columns;
      if (!colOmitted) {
        const s = ctx.scalar(colsVal!);
        if (isError(s)) return s;
        const n = ctx.number(s);
        if (!Number.isFinite(n)) return error('#VALUE!');
        colArg = Math.trunc(n);
      }
      if (rowArg === 0 || colArg === 0) return error('#CALC!');

      const rTake = Math.min(Math.abs(rowArg), arr.rows);
      const cTake = Math.min(Math.abs(colArg), arr.columns);
      const rStart = rowArg < 0 ? arr.rows - rTake : 0;
      const cStart = colArg < 0 ? arr.columns - cTake : 0;
      const cells: Scalar[] = [];
      for (let r = 0; r < rTake; r++) for (let c = 0; c < cTake; c++) {
        cells.push(arr.at(rStart + r, cStart + c));
      }
      return arrayFromMatrix(rTake, cTake, cells);
    },
  },
  {
    name: 'DROP',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 3,
    signature: [['array', 'rows', '[columns]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const arr = evalArray(args[0]!, ctx);
      if (isError(arr)) return arr;
      const rowsVal = ctx.evaluate(args[1]!);
      const rowOmitted = isOmitted(rowsVal);
      const colsVal = args[2] ? ctx.evaluate(args[2]) : undefined;
      const colOmitted = !colsVal || isOmitted(colsVal);
      if (rowOmitted && colOmitted) return error('#VALUE!');

      let rowArg = 0;
      if (!rowOmitted) {
        const s = ctx.scalar(rowsVal);
        if (isError(s)) return s;
        const n = ctx.number(s);
        if (!Number.isFinite(n)) return error('#VALUE!');
        rowArg = Math.trunc(n);
      }
      let colArg = 0;
      if (!colOmitted) {
        const s = ctx.scalar(colsVal!);
        if (isError(s)) return s;
        const n = ctx.number(s);
        if (!Number.isFinite(n)) return error('#VALUE!');
        colArg = Math.trunc(n);
      }

      const rRemain = arr.rows - Math.abs(rowArg);
      const cRemain = arr.columns - Math.abs(colArg);
      if (rRemain <= 0 || cRemain <= 0) return error('#CALC!');
      const rStart = rowArg > 0 ? rowArg : 0;
      const cStart = colArg > 0 ? colArg : 0;
      const cells: Scalar[] = [];
      for (let r = 0; r < rRemain; r++) for (let c = 0; c < cRemain; c++) {
        cells.push(arr.at(rStart + r, cStart + c));
      }
      return arrayFromMatrix(rRemain, cRemain, cells);
    },
  },
  {
    name: 'CHOOSEROWS',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 255,
    signature: [['array', 'row_num1', '[row_num2]'], true],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2) return error('#VALUE!');
      const arr = evalArray(args[0]!, ctx);
      if (isError(arr)) return arr;
      const picked: number[] = [];
      for (let i = 1; i < args.length; i++) {
        const list = readIntList(args[i]!, ctx);
        if (isError(list)) return list;
        for (const rawIdx of list) {
          if (rawIdx === 0 || Math.abs(rawIdx) > arr.rows) return error('#VALUE!');
          picked.push(rawIdx > 0 ? rawIdx - 1 : arr.rows + rawIdx);
        }
      }
      const cells: Scalar[] = [];
      for (const r of picked) for (let c = 0; c < arr.columns; c++) cells.push(arr.at(r, c));
      return arrayFromMatrix(picked.length, arr.columns, cells);
    },
  },
  {
    name: 'CHOOSECOLS',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 255,
    signature: [['array', 'col_num1', '[col_num2]'], true],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2) return error('#VALUE!');
      const arr = evalArray(args[0]!, ctx);
      if (isError(arr)) return arr;
      const picked: number[] = [];
      for (let i = 1; i < args.length; i++) {
        const list = readIntList(args[i]!, ctx);
        if (isError(list)) return list;
        for (const rawIdx of list) {
          if (rawIdx === 0 || Math.abs(rawIdx) > arr.columns) return error('#VALUE!');
          picked.push(rawIdx > 0 ? rawIdx - 1 : arr.columns + rawIdx);
        }
      }
      const cells: Scalar[] = [];
      for (let r = 0; r < arr.rows; r++) for (const c of picked) cells.push(arr.at(r, c));
      return arrayFromMatrix(arr.rows, picked.length, cells);
    },
  },
  {
    name: 'EXPAND',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 4,
    signature: [['array', 'rows', '[columns]', '[pad_with]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2 || args.length > 4) return error('#VALUE!');
      const arr = evalArray(args[0]!, ctx);
      if (isError(arr)) return arr;
      const rows = intArg(args[1], arr.rows, ctx);
      if (isError(rows)) return rows;
      const columns = intArg(args[2], arr.columns, ctx);
      if (isError(columns)) return columns;
      if (rows < arr.rows || columns < arr.columns) return error('#VALUE!');
      if (rows * columns > 100_000) return error('#NUM!');
      let pad: Scalar = error('#N/A');
      if (args[3]) {
        const padVal = ctx.evaluate(args[3]);
        if (!isOmitted(padVal)) pad = ctx.scalar(padVal);
      }
      const cells: Scalar[] = new Array(rows * columns);
      for (let r = 0; r < rows; r++) for (let c = 0; c < columns; c++) {
        cells[r * columns + c] = r < arr.rows && c < arr.columns ? arr.at(r, c) : pad;
      }
      return arrayFromMatrix(rows, columns, cells);
    },
  },
  {
    name: 'TEXTSPLIT',
    category: 'dynamic-array',
    minArgs: 2,
    maxArgs: 6,
    signature: [['text', '[col_delimiter]', '[row_delimiter]', '[ignore_empty]', '[match_mode]', '[pad_with]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2 || args.length > 6) return error('#VALUE!');
      const textScalar = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(textScalar)) return textScalar;
      const source = ctx.text(textScalar);
      const colDelims = readDelimiters(args[1], ctx);
      if (isError(colDelims)) return colDelims;
      const rowDelims = readDelimiters(args[2], ctx);
      if (isError(rowDelims)) return rowDelims;
      if (!colDelims.length && !rowDelims.length) return error('#VALUE!');

      const ignoreEmpty = boolArg(args[3], false, ctx);
      if (isError(ignoreEmpty)) return ignoreEmpty;
      const matchMode = intArg(args[4], 0, ctx);
      if (isError(matchMode)) return matchMode;
      if (matchMode !== 0 && matchMode !== 1) return error('#VALUE!');
      let pad: Scalar = error('#N/A');
      if (args[5]) {
        const padVal = ctx.evaluate(args[5]);
        if (!isOmitted(padVal)) pad = ctx.scalar(padVal);
      }

      const caseInsensitive = matchMode === 1;
      let rows = rowDelims.length ? splitMulti(source, rowDelims, caseInsensitive) : [source];
      if (ignoreEmpty) rows = rows.filter(r => r.length > 0);
      if (!rows.length) return error('#CALC!');

      const grid: string[][] = [];
      let maxCols = 0;
      for (const rowText of rows) {
        let parts = colDelims.length ? splitMulti(rowText, colDelims, caseInsensitive) : [rowText];
        if (ignoreEmpty) parts = parts.filter(p => p.length > 0);
        if (!parts.length && ignoreEmpty) continue;
        grid.push(parts);
        maxCols = Math.max(maxCols, parts.length);
      }
      if (!grid.length || maxCols === 0) return error('#CALC!');
      const cells: Scalar[] = new Array(grid.length * maxCols);
      for (let r = 0; r < grid.length; r++) {
        const line = grid[r]!;
        for (let c = 0; c < maxCols; c++) {
          cells[r * maxCols + c] = c < line.length ? line[c]! : pad;
        }
      }
      return arrayFromMatrix(grid.length, maxCols, cells);
    },
  },
];
