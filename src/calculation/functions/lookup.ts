import { columnLetters, type FormulaExpression } from '../../core/formula-references';
import { documentLimits } from '../../core/workbook';
import {
  arrayFromMatrix, compare, error, isArray, isError, isOmitted, isRange,
  type FormulaError, type FormulaValue, type Range, type Scalar,
} from '../formula-values';
import { logicalValue, wildcard } from './helpers';
import type { FunctionContext, FunctionSpec } from './types';

interface TableSource {
  readonly range?: Range;
  readonly rows: number;
  readonly columns: number;
  readonly at: (row: number, column: number) => Scalar;
}

function asTable(value: FormulaValue, ctx: FunctionContext): TableSource | FormulaError {
  if (isError(value)) return value;
  if (isRange(value)) {
    if (value.length !== 1) return error('#VALUE!');
    const area = value[0]!;
    return {
      range: area,
      rows: area.bottom - area.top + 1,
      columns: area.right - area.left + 1,
      at: (r, c) => ctx.read(area, area.top + r, area.left + c),
    };
  }
  if (isArray(value)) {
    return { rows: value.rows, columns: value.columns, at: (r, c) => value.at(r, c) };
  }
  const scalar = ctx.scalar(value);
  if (isError(scalar)) return scalar;
  return { rows: 1, columns: 1, at: () => scalar };
}

function gridLookup(vertical: boolean, args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
  if (args.length < 3 || args.length > 4) return error('#VALUE!');
  const key = ctx.scalar(ctx.evaluate(args[0]!));
  const tableVal = ctx.evaluate(args[1]!, true);
  const indexVal = ctx.scalar(ctx.evaluate(args[2]!));
  const modeVal = args[3] ? ctx.scalar(ctx.evaluate(args[3])) : true;
  if (isError(key)) return key;
  if (isError(indexVal)) return indexVal;
  if (isError(modeVal)) return modeVal;
  if (!isRange(tableVal) && !isArray(tableVal)) return isError(tableVal) ? tableVal : error('#VALUE!');
  if (isRange(tableVal) && tableVal.length !== 1) return error('#VALUE!');
  const table = asTable(tableVal, ctx);
  if (isError(table)) return table;
  const offset = Math.trunc(ctx.number(indexVal)) - 1;
  const approximate = logicalValue(modeVal);
  if (Number.isNaN(offset) || offset < 0 || approximate === undefined) return error('#VALUE!');
  if (offset >= (vertical ? table.columns : table.rows)) return error('#REF!');
  if (key === null) return error('#N/A');

  const lineLen = vertical ? table.rows : table.columns;
  const headerAt = (i: number) => vertical ? table.at(i, 0) : table.at(0, i);
  const resultAt = (i: number) => vertical ? table.at(i, offset) : table.at(offset, i);

  if (vertical && table.range) {
    const area = table.range;
    let foundRow: number | undefined;
    for (const entry of ctx.entries(area)) {
      if (entry.column !== area.left) continue;
      const candidate = entry.value;
      if (candidate === null || isError(candidate)) continue;
      if (approximate && typeof key !== typeof candidate) continue;
      const comparison = compare(candidate, key);
      if (!approximate && comparison === 0) return table.at(entry.row - area.top, offset);
      if (approximate && comparison <= 0) foundRow = entry.row;
      if (approximate && comparison > 0) break;
    }
    return foundRow === undefined ? error('#N/A') : table.at(foundRow - area.top, offset);
  }

  let found: number | undefined;
  for (let i = 0; i < lineLen; i++) {
    const candidate = headerAt(i);
    if (candidate === null || isError(candidate)) continue;
    if (approximate && typeof key !== typeof candidate) continue;
    const comparison = compare(candidate, key);
    if (!approximate && comparison === 0) return resultAt(i);
    if (approximate && comparison <= 0) found = i;
    if (approximate && comparison > 0) break;
  }
  return found === undefined ? error('#N/A') : resultAt(found);
}

function sparseEntries(table: TableSource, horizontal: boolean, ctx: FunctionContext): { index: number; value: Scalar }[] | undefined {
  if (!table.range) return undefined;
  const area = table.range;
  const list: { index: number; value: Scalar }[] = [];
  for (const entry of ctx.entries(area)) {
    list.push({ index: horizontal ? entry.column - area.left : entry.row - area.top, value: entry.value });
  }
  return list;
}

function findFirstNullIndex(populated: readonly { index: number; value: Scalar }[], length: number, reverse: boolean): number | undefined {
  if (!reverse) {
    let expected = 0;
    for (const item of populated) {
      if (item.index > expected) return expected;
      if (item.value === null) return item.index;
      expected = item.index + 1;
    }
    return expected < length ? expected : undefined;
  }
  const byIdx = new Map(populated.map(p => [p.index, p.value]));
  for (let i = length - 1; i >= 0; i--) {
    const val = byIdx.get(i);
    if (val === undefined || val === null) return i;
  }
  return undefined;
}

function findXIndex(
  key: Exclude<Scalar, FormulaError>,
  readAt: (i: number) => Scalar,
  length: number,
  matchMode: number,
  searchMode: number,
  ctx: FunctionContext,
  sparse?: readonly { index: number; value: Scalar }[],
): number | FormulaError {
  if (![0, -1, 1, 2].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) return error('#VALUE!');
  if (matchMode === 2 && Math.abs(searchMode) === 2) return error('#VALUE!');
  if (length <= 0) return error('#N/A');

  if (sparse && length > 1000) {
    if (matchMode === 2) {
      const matcher = wildcard(ctx.text(key));
      const items = searchMode === -1 ? [...sparse].reverse() : sparse;
      for (const item of items) {
        if (typeof item.value === 'string' && matcher(item.value)) return item.index;
      }
      return error('#N/A');
    }
    if (matchMode === 0 && key === null) {
      return findFirstNullIndex(sparse, length, searchMode === -1) ?? error('#N/A');
    }
    const items = searchMode === -1 || searchMode === -2 ? [...sparse].reverse() : sparse;
    let bestIdx: number | undefined;
    let bestVal: Exclude<Scalar, FormulaError | null> | undefined;
    for (const item of items) {
      const cand = item.value;
      if (cand === null || isError(cand)) continue;
      const cmp = compare(cand, key);
      if (cmp === 0) return item.index;
      if (matchMode === 0 || typeof cand !== typeof key) continue;
      if (matchMode === -1 && cmp < 0 && (bestVal === undefined || compare(cand, bestVal) > 0)) {
        bestVal = cand;
        bestIdx = item.index;
      } else if (matchMode === 1 && cmp > 0 && (bestVal === undefined || compare(cand, bestVal) < 0)) {
        bestVal = cand;
        bestIdx = item.index;
      }
    }
    return bestIdx === undefined ? error('#N/A') : bestIdx;
  }

  if (matchMode === 2) {
    const matcher = wildcard(ctx.text(key));
    const step = searchMode === -1 ? -1 : 1;
    const start = searchMode === -1 ? length - 1 : 0;
    for (let k = 0, i = start; k < length; k++, i += step) {
      const cand = readAt(i);
      if (typeof cand === 'string' && matcher(cand)) return i;
    }
    return error('#N/A');
  }

  if (Math.abs(searchMode) === 2) {
    const ascending = searchMode === 2;
    let lo = 0, hi = length - 1;
    let bestExact: number | undefined;
    let bestApprox: number | undefined;
    while (lo <= hi) {
      const mid = (lo + hi) >>> 1;
      const cand = readAt(mid);
      if (cand === null || isError(cand)) {
        // Skip empty slots by probing inward; binary search assumes contiguous sorted values.
        lo = mid + 1;
        continue;
      }
      const cmp = compare(cand, key);
      if (cmp === 0) {
        bestExact = mid;
        break;
      }
      if (matchMode === -1 && cmp < 0) bestApprox = mid;
      if (matchMode === 1 && cmp > 0) bestApprox = mid;
      if (ascending ? cmp < 0 : cmp > 0) lo = mid + 1;
      else hi = mid - 1;
    }
    if (bestExact !== undefined) return bestExact;
    return bestApprox === undefined ? error('#N/A') : bestApprox;
  }

  const step = searchMode === -1 ? -1 : 1;
  const start = searchMode === -1 ? length - 1 : 0;
  let bestIdx: number | undefined;
  let bestVal: Exclude<Scalar, FormulaError | null> | undefined;

  for (let k = 0, i = start; k < length; k++, i += step) {
    const cand = readAt(i);
    if (isError(cand)) continue;
    if (key === null) {
      if (cand === null) return i;
      continue;
    }
    if (cand === null) continue;
    const cmp = compare(cand, key);
    if (cmp === 0) return i;
    if (matchMode === 0) continue;
    if (typeof cand !== typeof key) continue;
    if (matchMode === -1 && cmp < 0) {
      if (bestVal === undefined || compare(cand, bestVal) > 0) {
        bestVal = cand;
        bestIdx = i;
      }
    } else if (matchMode === 1 && cmp > 0) {
      if (bestVal === undefined || compare(cand, bestVal) < 0) {
        bestVal = cand;
        bestIdx = i;
      }
    }
  }
  return bestIdx === undefined ? error('#N/A') : bestIdx;
}

function intArg(expr: FormulaExpression | undefined, fallback: number, ctx: FunctionContext): number | FormulaError {
  if (!expr) return fallback;
  const raw = ctx.evaluate(expr);
  if (isOmitted(raw)) return fallback;
  const scalar = ctx.scalar(raw);
  if (isError(scalar)) return scalar;
  const n = ctx.number(scalar);
  return Number.isFinite(n) ? Math.trunc(n) : error('#VALUE!');
}

function sliceRange(area: Range, r0: number, r1: number, c0: number, c1: number): Range {
  return { sheetId: area.sheetId, top: area.top + r0, bottom: area.top + r1, left: area.left + c0, right: area.left + c1 };
}

function formatSheetPrefix(sheetName: string): string {
  if (!sheetName) return '';
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)
    ? `${sheetName}!`
    : `'${sheetName.replaceAll("'", "''")}'!`;
}

export const lookupFunctions: readonly FunctionSpec[] = [
  {
    name: 'VLOOKUP',
    category: 'lookup',
    minArgs: 3,
    maxArgs: 4,
    signature: [['lookup_value', 'table_array', 'col_index_num', '[range_lookup]']],
    evaluate: (args, ctx) => gridLookup(true, args, ctx),
  },
  {
    name: 'HLOOKUP',
    category: 'lookup',
    minArgs: 3,
    maxArgs: 4,
    signature: [['lookup_value', 'table_array', 'row_index_num', '[range_lookup]']],
    evaluate: (args, ctx) => gridLookup(false, args, ctx),
  },
  {
    name: 'XLOOKUP',
    category: 'lookup',
    minArgs: 3,
    maxArgs: 6,
    signature: [['lookup_value', 'lookup_array', 'return_array', '[if_not_found]', '[match_mode]', '[search_mode]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 3 || args.length > 6) return error('#VALUE!');
      const key = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(key)) return key;
      const lookupTable = asTable(ctx.evaluate(args[1]!, true), ctx);
      if (isError(lookupTable)) return lookupTable;
      if (lookupTable.rows > 1 && lookupTable.columns > 1) return error('#VALUE!');
      const returnTable = asTable(ctx.evaluate(args[2]!, true), ctx);
      if (isError(returnTable)) return returnTable;
      const matchMode = intArg(args[4], 0, ctx);
      if (isError(matchMode)) return matchMode;
      const searchMode = intArg(args[5], 1, ctx);
      if (isError(searchMode)) return searchMode;

      const horizontal = lookupTable.rows === 1 && lookupTable.columns > 1;
      const lookupLen = horizontal ? lookupTable.columns : lookupTable.rows;
      const returnLen = horizontal ? returnTable.columns : returnTable.rows;
      if (lookupLen !== returnLen) return error('#VALUE!');

      const idx = findXIndex(
        key,
        i => horizontal ? lookupTable.at(0, i) : lookupTable.at(i, 0),
        lookupLen,
        matchMode,
        searchMode,
        ctx,
        sparseEntries(lookupTable, horizontal, ctx),
      );
      if (isError(idx)) {
        if (idx.value === '#N/A' && args[3]) {
          const fallback = ctx.evaluate(args[3], true);
          if (!isOmitted(fallback)) return fallback;
        }
        return idx;
      }

      if (returnTable.range) {
        return [horizontal
          ? sliceRange(returnTable.range, 0, returnTable.rows - 1, idx, idx)
          : sliceRange(returnTable.range, idx, idx, 0, returnTable.columns - 1)];
      }
      if (horizontal) {
        if (returnTable.rows === 1) return returnTable.at(0, idx);
        return arrayFromMatrix(returnTable.rows, 1, Array.from({ length: returnTable.rows }, (_, r) => returnTable.at(r, idx)));
      }
      if (returnTable.columns === 1) return returnTable.at(idx, 0);
      return arrayFromMatrix(1, returnTable.columns, Array.from({ length: returnTable.columns }, (_, c) => returnTable.at(idx, c)));
    },
  },
  {
    name: 'XMATCH',
    category: 'lookup',
    minArgs: 2,
    maxArgs: 4,
    signature: [['lookup_value', 'lookup_array', '[match_mode]', '[search_mode]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 4) return error('#VALUE!');
      const key = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(key)) return key;
      const lookupTable = asTable(ctx.evaluate(args[1]!, true), ctx);
      if (isError(lookupTable)) return lookupTable;
      if (lookupTable.rows > 1 && lookupTable.columns > 1) return error('#VALUE!');
      const matchMode = intArg(args[2], 0, ctx);
      if (isError(matchMode)) return matchMode;
      const searchMode = intArg(args[3], 1, ctx);
      if (isError(searchMode)) return searchMode;
      const horizontal = lookupTable.rows === 1 && lookupTable.columns > 1;
      const len = horizontal ? lookupTable.columns : lookupTable.rows;
      const idx = findXIndex(
        key,
        i => horizontal ? lookupTable.at(0, i) : lookupTable.at(i, 0),
        len,
        matchMode,
        searchMode,
        ctx,
        sparseEntries(lookupTable, horizontal, ctx),
      );
      return isError(idx) ? idx : idx + 1;
    },
  },
  {
    name: 'MATCH',
    category: 'lookup',
    minArgs: 2,
    maxArgs: 3,
    signature: [['lookup_value', 'lookup_array', '[match_type]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const key = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(key)) return key;
      const arrayVal = ctx.evaluate(args[1]!, true);
      if (!isRange(arrayVal) && !isArray(arrayVal)) return isError(arrayVal) ? arrayVal : error('#N/A');
      const table = asTable(arrayVal, ctx);
      if (isError(table)) return table;
      if (table.rows > 1 && table.columns > 1) return error('#N/A');
      const matchType = intArg(args[2], 1, ctx);
      if (isError(matchType)) return matchType;

      const horizontal = table.rows === 1 && table.columns > 1;
      const len = horizontal ? table.columns : table.rows;
      const sparse = len > 1000 ? sparseEntries(table, horizontal, ctx) : undefined;
      const readAt = (i: number) => horizontal ? table.at(0, i) : table.at(i, 0);

      if (matchType === 0) {
        if (typeof key === 'string' && /[*?~]/u.test(key)) {
          const matcher = wildcard(key);
          if (sparse) {
            for (const item of sparse) {
              if (typeof item.value === 'string' && matcher(item.value)) return item.index + 1;
            }
            return error('#N/A');
          }
          for (let i = 0; i < len; i++) {
            const cand = readAt(i);
            if (typeof cand === 'string' && matcher(cand)) return i + 1;
          }
          return error('#N/A');
        }
        if (key === null) return error('#N/A');
        if (sparse) {
          for (const item of sparse) {
            if (item.value === null || isError(item.value)) continue;
            if (compare(item.value, key) === 0) return item.index + 1;
          }
          return error('#N/A');
        }
        for (let i = 0; i < len; i++) {
          const cand = readAt(i);
          if (cand === null || isError(cand)) continue;
          if (compare(cand, key) === 0) return i + 1;
        }
        return error('#N/A');
      }

      if (key === null) return error('#N/A');
      const ascending = matchType > 0;
      let best: number | undefined;
      if (sparse) {
        for (const item of sparse) {
          const cand = item.value;
          if (cand === null || isError(cand) || typeof cand !== typeof key) continue;
          const cmp = compare(cand, key);
          if (ascending ? cmp <= 0 : cmp >= 0) best = item.index;
          else break;
        }
        return best === undefined ? error('#N/A') : best + 1;
      }
      for (let i = 0; i < len; i++) {
        const cand = readAt(i);
        if (cand === null || isError(cand) || typeof cand !== typeof key) continue;
        const cmp = compare(cand, key);
        if (ascending) {
          if (cmp <= 0) best = i;
          else break;
        } else {
          if (cmp >= 0) best = i;
          else break;
        }
      }
      return best === undefined ? error('#N/A') : best + 1;
    },
  },
  {
    name: 'LOOKUP',
    category: 'lookup',
    minArgs: 2,
    maxArgs: 3,
    signature: [['lookup_value', 'lookup_vector', '[result_vector]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const key = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(key)) return key;
      const firstVal = ctx.evaluate(args[1]!, true);
      const first = asTable(firstVal, ctx);
      if (isError(first)) return first;
      if (key === null) return error('#N/A');

      let searchLen: number;
      let readKey: (i: number) => Scalar;
      let readResult: (i: number) => Scalar;

      if (args[2]) {
        if (first.rows > 1 && first.columns > 1) return error('#VALUE!');
        const second = asTable(ctx.evaluate(args[2], true), ctx);
        if (isError(second)) return second;
        if (second.rows > 1 && second.columns > 1) return error('#VALUE!');
        const firstHoriz = first.rows === 1 && first.columns > 1;
        const secondHoriz = second.rows === 1 && second.columns > 1;
        searchLen = firstHoriz ? first.columns : first.rows;
        const secondLen = secondHoriz ? second.columns : second.rows;
        readKey = i => firstHoriz ? first.at(0, i) : first.at(i, 0);
        readResult = i => i < secondLen ? (secondHoriz ? second.at(0, i) : second.at(i, 0)) : error('#N/A');
      } else {
        const useRows = first.rows >= first.columns;
        searchLen = useRows ? first.rows : first.columns;
        readKey = i => useRows ? first.at(i, 0) : first.at(0, i);
        readResult = i => useRows ? first.at(i, first.columns - 1) : first.at(first.rows - 1, i);
      }

      let found: number | undefined;
      if (searchLen > 1000 && first.range) {
        const horiz = !args[2] ? first.rows < first.columns : first.rows === 1 && first.columns > 1;
        for (const entry of ctx.entries(first.range)) {
          if (!args[2] && (horiz ? entry.row !== first.range.top : entry.column !== first.range.left)) continue;
          const idx = horiz ? entry.column - first.range.left : entry.row - first.range.top;
          const cand = entry.value;
          if (cand === null || isError(cand) || typeof cand !== typeof key) continue;
          const cmp = compare(cand, key);
          if (cmp <= 0) found = idx;
          else break;
        }
        return found === undefined ? error('#N/A') : readResult(found);
      }
      for (let i = 0; i < searchLen; i++) {
        const cand = readKey(i);
        if (cand === null || isError(cand) || typeof cand !== typeof key) continue;
        const cmp = compare(cand, key);
        if (cmp <= 0) found = i;
        else break;
      }
      return found === undefined ? error('#N/A') : readResult(found);
    },
  },
  {
    name: 'INDEX',
    category: 'lookup',
    minArgs: 2,
    maxArgs: 4,
    signature: [['array', 'row_num', '[column_num]', '[area_num]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2 || args.length > 4) return error('#VALUE!');
      const targetVal = ctx.evaluate(args[0]!, true);
      if (isError(targetVal)) return targetVal;
      const areaNum = intArg(args[3], 1, ctx);
      if (isError(areaNum)) return areaNum;

      let table: TableSource;
      if (isRange(targetVal)) {
        if (areaNum < 1 || areaNum > targetVal.length) return error('#REF!');
        const area = targetVal[areaNum - 1]!;
        table = {
          range: area,
          rows: area.bottom - area.top + 1,
          columns: area.right - area.left + 1,
          at: (r, c) => ctx.read(area, area.top + r, area.left + c),
        };
      } else {
        if (areaNum !== 1) return error('#REF!');
        const converted = asTable(targetVal, ctx);
        if (isError(converted)) return converted;
        table = converted;
      }

      const rawRow = intArg(args[1], 0, ctx);
      if (isError(rawRow)) return rawRow;
      let rowNum: number;
      let colNum: number;

      if (args.length === 2) {
        if (table.rows === 1 && table.columns > 1) {
          rowNum = 1;
          colNum = rawRow;
        } else if (table.columns === 1) {
          rowNum = rawRow;
          colNum = 1;
        } else {
          rowNum = rawRow;
          colNum = 0;
        }
      } else {
        const rawCol = intArg(args[2], 0, ctx);
        if (isError(rawCol)) return rawCol;
        rowNum = rawRow;
        colNum = rawCol;
      }

      if (rowNum < 0 || colNum < 0) return error('#VALUE!');
      if (rowNum > table.rows || colNum > table.columns) return error('#REF!');

      if (table.range) {
        const r0 = rowNum === 0 ? 0 : rowNum - 1;
        const r1 = rowNum === 0 ? table.rows - 1 : rowNum - 1;
        const c0 = colNum === 0 ? 0 : colNum - 1;
        const c1 = colNum === 0 ? table.columns - 1 : colNum - 1;
        return [sliceRange(table.range, r0, r1, c0, c1)];
      }
      if (rowNum > 0 && colNum > 0) return table.at(rowNum - 1, colNum - 1);
      const r0 = rowNum === 0 ? 0 : rowNum - 1;
      const outRows = rowNum === 0 ? table.rows : 1;
      const c0 = colNum === 0 ? 0 : colNum - 1;
      const outCols = colNum === 0 ? table.columns : 1;
      const cells: Scalar[] = [];
      for (let r = 0; r < outRows; r++) for (let c = 0; c < outCols; c++) cells.push(table.at(r0 + r, c0 + c));
      return arrayFromMatrix(outRows, outCols, cells);
    },
  },
  {
    name: 'CHOOSE',
    category: 'lookup',
    minArgs: 2,
    maxArgs: 255,
    signature: [['index_num', 'value1', '[value2]'], true],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 2) return error('#VALUE!');
      const raw = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(raw)) return raw;
      const index = Math.trunc(ctx.number(raw));
      if (!Number.isFinite(index) || index < 1 || index >= args.length) return error('#VALUE!');
      return ctx.evaluate(args[index]!, true);
    },
  },
  {
    name: 'OFFSET',
    category: 'lookup',
    minArgs: 3,
    maxArgs: 5,
    volatile: true,
    signature: [['reference', 'rows', 'cols', '[height]', '[width]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 3 || args.length > 5) return error('#VALUE!');
      const refVal = ctx.evaluate(args[0]!, true);
      if (isError(refVal)) return refVal;
      if (!isRange(refVal) || refVal.length !== 1) return error('#VALUE!');
      const base = refVal[0]!;
      const rowOff = intArg(args[1], 0, ctx);
      if (isError(rowOff)) return rowOff;
      const colOff = intArg(args[2], 0, ctx);
      if (isError(colOff)) return colOff;
      const baseH = base.bottom - base.top + 1;
      const baseW = base.right - base.left + 1;
      const height = intArg(args[3], baseH, ctx);
      if (isError(height)) return height;
      const width = intArg(args[4], baseW, ctx);
      if (isError(width)) return width;
      if (height <= 0 || width <= 0) return error('#REF!');
      const top = base.top + rowOff;
      const left = base.left + colOff;
      const bottom = top + height - 1;
      const right = left + width - 1;
      if (top < 1 || left < 1 || bottom > documentLimits.rows || right > documentLimits.columns) return error('#REF!');
      return [{ sheetId: base.sheetId, top, left, bottom, right }];
    },
  },
  {
    name: 'INDIRECT',
    category: 'lookup',
    minArgs: 1,
    maxArgs: 2,
    volatile: true,
    signature: [['ref_text', '[a1]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length < 1 || args.length > 2) return error('#VALUE!');
      const refVal = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(refVal)) return refVal;
      if (typeof refVal !== 'string') return error('#REF!');
      let a1 = true;
      if (args[1]) {
        const modeVal = ctx.evaluate(args[1]);
        if (!isOmitted(modeVal)) {
          const modeScalar = ctx.scalar(modeVal);
          if (isError(modeScalar)) return modeScalar;
          const bool = logicalValue(modeScalar);
          if (bool === undefined) return error('#VALUE!');
          a1 = bool;
        }
      }
      return ctx.resolveRefText(refVal, a1) ?? error('#REF!');
    },
  },
  {
    name: 'ROW',
    category: 'lookup',
    minArgs: 0,
    maxArgs: 1,
    signature: [['[reference]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length > 1) return error('#VALUE!');
      if (args.length === 0) return ctx.target.row;
      const val = ctx.evaluate(args[0]!, true);
      if (isOmitted(val)) return ctx.target.row;
      if (isError(val)) return val;
      if (!isRange(val) || val.length !== 1) return error('#VALUE!');
      const area = val[0]!;
      const rows = area.bottom - area.top + 1;
      if (rows === 1) return area.top;
      return { kind: 'array', rows, columns: 1, at: (r: number) => area.top + r };
    },
  },
  {
    name: 'ROWS',
    category: 'lookup',
    minArgs: 1,
    maxArgs: 1,
    signature: [['array']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 1) return error('#VALUE!');
      const val = ctx.evaluate(args[0]!, true);
      if (isError(val)) return val;
      if (isRange(val)) return val.length === 1 ? val[0]!.bottom - val[0]!.top + 1 : error('#VALUE!');
      if (isArray(val)) return val.rows;
      return error('#VALUE!');
    },
  },
  {
    name: 'COLUMN',
    category: 'lookup',
    minArgs: 0,
    maxArgs: 1,
    signature: [['[reference]']],
    evaluate(args, ctx): FormulaValue {
      if (args.length > 1) return error('#VALUE!');
      if (args.length === 0) return ctx.target.column;
      const val = ctx.evaluate(args[0]!, true);
      if (isOmitted(val)) return ctx.target.column;
      if (isError(val)) return val;
      if (!isRange(val) || val.length !== 1) return error('#VALUE!');
      const area = val[0]!;
      const columns = area.right - area.left + 1;
      if (columns === 1) return area.left;
      return { kind: 'array', rows: 1, columns, at: (_: number, c: number) => area.left + c };
    },
  },
  {
    name: 'COLUMNS',
    category: 'lookup',
    minArgs: 1,
    maxArgs: 1,
    signature: [['array']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 1) return error('#VALUE!');
      const val = ctx.evaluate(args[0]!, true);
      if (isError(val)) return val;
      if (isRange(val)) return val.length === 1 ? val[0]!.right - val[0]!.left + 1 : error('#VALUE!');
      if (isArray(val)) return val.columns;
      return error('#VALUE!');
    },
  },
  {
    name: 'ADDRESS',
    category: 'lookup',
    minArgs: 2,
    maxArgs: 5,
    signature: [['row_num', 'column_num', '[abs_num]', '[a1]', '[sheet_text]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 5) return error('#VALUE!');
      const rowNum = intArg(args[0], 0, ctx);
      const colNum = intArg(args[1], 0, ctx);
      if (isError(rowNum)) return rowNum;
      if (isError(colNum)) return colNum;
      if (rowNum < 1 || rowNum > documentLimits.rows || colNum < 1 || colNum > documentLimits.columns) return error('#VALUE!');

      const absNum = intArg(args[2], 1, ctx);
      if (isError(absNum)) return absNum;
      if (absNum < 1 || absNum > 4) return error('#VALUE!');

      let a1 = true;
      if (args[3]) {
        const modeVal = ctx.evaluate(args[3]);
        if (!isOmitted(modeVal)) {
          const modeScalar = ctx.scalar(modeVal);
          if (isError(modeScalar)) return modeScalar;
          const bool = logicalValue(modeScalar);
          if (bool === undefined) return error('#VALUE!');
          a1 = bool;
        }
      }

      let sheetPrefix = '';
      if (args[4]) {
        const rawSheet = ctx.evaluate(args[4]);
        if (!isOmitted(rawSheet)) {
          const sheetScalar = ctx.scalar(rawSheet);
          if (isError(sheetScalar)) return sheetScalar;
          sheetPrefix = formatSheetPrefix(ctx.text(sheetScalar));
        }
      }

      const absRow = absNum === 1 || absNum === 2;
      const absCol = absNum === 1 || absNum === 3;
      if (a1) {
        return `${sheetPrefix}${absCol ? '$' : ''}${columnLetters(colNum)}${absRow ? '$' : ''}${rowNum}`;
      }
      const rPart = absRow ? `R${rowNum}` : `R[${rowNum}]`;
      const cPart = absCol ? `C${colNum}` : `C[${colNum}]`;
      return `${sheetPrefix}${rPart}${cPart}`;
    },
  },
  {
    name: 'AREAS',
    category: 'lookup',
    minArgs: 1,
    maxArgs: 1,
    signature: [['reference']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 1) return error('#VALUE!');
      const val = ctx.evaluate(args[0]!, true);
      if (isError(val)) return val;
      return isRange(val) ? val.length : error('#VALUE!');
    },
  },
  {
    name: 'FORMULATEXT',
    category: 'lookup',
    minArgs: 1,
    maxArgs: 1,
    signature: [['reference']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 1) return error('#VALUE!');
      const val = ctx.evaluate(args[0]!, true);
      if (isError(val)) return val;
      if (!isRange(val) || val.length !== 1) return error('#VALUE!');
      const area = val[0]!;
      const cell = ctx.cellAt(area.sheetId, area.top, area.left);
      return cell?.content.kind === 'formula' ? cell.content.source : error('#N/A');
    },
  },
  {
    name: 'HYPERLINK',
    category: 'lookup',
    minArgs: 1,
    maxArgs: 2,
    signature: [['link_location', '[friendly_name]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 2) return error('#VALUE!');
      const link = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(link)) return link;
      if (!args[1]) return link === null ? '' : link;
      const friendly = ctx.evaluate(args[1]);
      if (isOmitted(friendly)) return link === null ? '' : link;
      const label = ctx.scalar(friendly);
      return isError(label) ? label : label === null ? '' : label;
    },
  },
];
