import type { FormulaExpression } from '../../core/formula-references';
import { compare, error, isArray, isError, isRange, type Scalar } from '../formula-values';
import { logicalValue } from './helpers';
import type { FunctionContext, FunctionSpec } from './types';

function vlookup(args: readonly FormulaExpression[], ctx: FunctionContext): Scalar {
  if (args.length < 3 || args.length > 4) return error('#VALUE!');
  const key = ctx.scalar(ctx.evaluate(args[0]!));
  const table = ctx.evaluate(args[1]!, true);
  const index = ctx.scalar(ctx.evaluate(args[2]!));
  const mode = args[3] ? ctx.scalar(ctx.evaluate(args[3])) : true;
  if (isError(key)) return key;
  if (isError(index)) return index;
  if (isError(mode)) return mode;
  if (!isRange(table) && !isArray(table)) return isError(table) ? table : error('#VALUE!');
  if (isRange(table) && table.length !== 1) return error('#VALUE!');
  const range = isRange(table) ? table[0]! : undefined;
  const rows = isArray(table) ? table.rows : range!.bottom - range!.top + 1;
  const columns = isArray(table) ? table.columns : range!.right - range!.left + 1;
  const at = (row: number, column: number) => isArray(table)
    ? table.at(row, column)
    : ctx.read(range!, range!.top + row, range!.left + column);
  const offset = Math.trunc(ctx.number(index)) - 1;
  const approximate = logicalValue(mode);
  if (Number.isNaN(offset) || offset < 0 || approximate === undefined) return error('#VALUE!');
  if (offset >= columns) return error('#REF!');
  if (key === null) return error('#N/A');
  if (range) {
    let foundRow: number | undefined;
    for (const entry of ctx.entries(range)) {
      if (entry.column !== range.left) continue;
      const candidate = entry.value;
      if (candidate === null || isError(candidate)) continue;
      if (approximate && typeof key !== typeof candidate) continue;
      const comparison = compare(candidate, key);
      if (!approximate && comparison === 0) return at(entry.row - range.top, offset);
      if (approximate && comparison <= 0) foundRow = entry.row;
      if (approximate && comparison > 0) break;
    }
    return foundRow === undefined ? error('#N/A') : at(foundRow - range.top, offset);
  }
  let found: number | undefined;
  for (let row = 0; row < rows; row++) {
    const candidate = at(row, 0);
    if (candidate === null || isError(candidate)) continue;
    if (approximate && typeof key !== typeof candidate) continue;
    const comparison = compare(candidate, key);
    if (!approximate && comparison === 0) {
      found = row;
      break;
    }
    if (approximate && comparison <= 0) found = row;
    if (approximate && comparison > 0) break;
  }
  return found === undefined ? error('#N/A') : at(found, offset);
}

export const lookupFunctions: readonly FunctionSpec[] = [
  {
    name: 'VLOOKUP',
    category: 'lookup',
    minArgs: 3,
    maxArgs: 4,
    signature: [['lookup_value', 'table_array', 'col_index_num', '[range_lookup]']],
    evaluate: vlookup,
  },
];
