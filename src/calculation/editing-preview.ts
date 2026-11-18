import type { CacheUpdate } from '../application/workbook-protocol';
import { parseFormula, referenceBox, sameSheet, type FormulaExpression } from '../core/formula-references';
import {
  cellKey, documentLimits, findCell, formulaErrors, type CellCache, type CellDocument, type SheetDocument, type WorkbookDocument,
} from '../core/workbook';

type Value = number | boolean | string | null | { readonly kind: 'error'; readonly value: typeof formulaErrors[number] };
type Result = Value | readonly Value[] | undefined;
const error = (value: typeof formulaErrors[number]): Extract<Value, object> => ({ kind: 'error', value });
const isError = (value: Value): value is Extract<Value, object> => typeof value === 'object' && value !== null;
const isRange = (value: Result): value is readonly Value[] => Array.isArray(value);
const finite = (value: number): Value => Number.isFinite(value) ? value : error('#NUM!');
const aggregateNames = new Set(['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN']);

function cacheValue(cache: CellCache): Value {
  return cache.kind === 'error' ? cache : cache.value;
}
function toCache(value: Result): CellCache | undefined {
  if (value === undefined || isRange(value)) return undefined;
  if (value === null) return { kind: 'number', value: 0 };
  if (isError(value)) return value;
  if (typeof value === 'number') return { kind: 'number', value };
  if (typeof value === 'boolean') return { kind: 'boolean', value };
  return value.length <= documentLimits.textLength && !value.includes('\0') ? { kind: 'text', value } : error('#VALUE!');
}

function numeric(value: Value): number {
  if (value === null) return 0;
  if (typeof value === 'number' || typeof value === 'boolean') return Number(value);
  if (typeof value === 'string' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/iu.test(value.trim())) return Number(value);
  return NaN;
}

/** Small, bounded calculation used by the editing slice until the native engine is integrated. */
export function editingPreview(document: WorkbookDocument): readonly CacheUpdate[] {
  const results = new Map<CellDocument, Result>();
  const active = new Set<CellDocument>();
  const updates: CacheUpdate[] = [];
  const maxDepth = 128;
  let steps = 75_000;

  function cellValue(sheet: SheetDocument, cell: CellDocument | undefined, depth: number): Result {
    if (!cell || cell.content.kind === 'blank') return null;
    if (cell.content.kind === 'python') return undefined;
    if (cell.content.kind !== 'formula') return cell.content.value;
    if (cell.requiresExternalCalculation) return cell.cache && cacheValue(cell.cache);
    if (results.has(cell)) return results.get(cell);
    if (active.has(cell)) return error('#REF!');
    if (depth > maxDepth || (steps -= Math.ceil(cell.content.source.length / 16)) <= 0) return undefined;
    let tree: FormulaExpression;
    try { tree = parseFormula(cell.content.source).expression; }
    catch { results.set(cell, undefined); return undefined; }
    active.add(cell);
    const value = evaluate(sheet, tree, depth + 1);
    active.delete(cell);
    if (steps > 0) results.set(cell, value);
    return value;
  }

  function evaluate(sheet: SheetDocument, expression: FormulaExpression, depth: number): Result {
    if (--steps < 0 || depth > maxDepth) return undefined;
    const inner = (node: FormulaExpression) => evaluate(sheet, node, depth + 1);
    switch (expression.kind) {
      case 'literal': return typeof expression.value === 'number' ? finite(expression.value) : expression.value;
      case 'error': return error(expression.value as typeof formulaErrors[number]);
      case 'name': return undefined;
      case 'reference': {
        const ref = expression.reference;
        const target = ref.sheet === undefined ? sheet : document.sheets.find(candidate => sameSheet(candidate.name, ref.sheet!));
        const { top, left, bottom, right } = referenceBox(ref);
        if (!target || bottom > documentLimits.rows || right > documentLimits.columns) return error('#REF!');
        if (!ref.range) return cellValue(target, findCell(target, { row: top, column: left }), depth + 1);
        const values: Value[] = [];
        const size = (bottom - top + 1) * (right - left + 1);
        if (size <= target.cells.length) {
          for (let row = top; row <= bottom; row++) for (let column = left; column <= right; column++) {
            if (--steps < 0) return undefined;
            const item = cellValue(target, findCell(target, { row, column }), depth + 1);
            if (item === undefined || isRange(item)) return undefined;
            values.push(item);
          }
        } else {
          if ((steps -= target.cells.length) < 0) return undefined;
          const matching = target.cells.filter(cell => cell.row >= top && cell.row <= bottom &&
            cell.column >= left && cell.column <= right).sort((a, b) => cellKey(a) - cellKey(b));
          for (const cell of matching) {
            const item = cellValue(target, cell, depth + 1);
            if (item === undefined || isRange(item)) return undefined;
            values.push(item);
          }
        }
        return values;
      }
      case 'union': {
        const values: Value[] = [];
        for (const part of expression.parts) {
          const item = inner(part);
          if (item === undefined) return undefined;
          if (isRange(item)) for (const member of item) values.push(member);
          else values.push(item);
        }
        return values;
      }
      case 'unary': {
        const value = inner(expression.value);
        if (value === undefined || isRange(value)) return undefined;
        if (isError(value)) return value;
        const number = numeric(value);
        return Number.isNaN(number) ? error('#VALUE!') :
          finite(expression.operator === '-' ? -number : expression.operator === '%' ? number / 100 : number);
      }
      case 'binary': {
        const left = inner(expression.left);
        if (left === undefined || isRange(left)) return undefined;
        if (isError(left)) return left;
        const right = inner(expression.right);
        if (right === undefined || isRange(right)) return undefined;
        if (isError(right)) return right;
        if (expression.operator === '&') return String(left ?? '') + String(right ?? '');
        const a = numeric(left), b = numeric(right);
        if (Number.isNaN(a) || Number.isNaN(b)) return error('#VALUE!');
        switch (expression.operator) {
          case '+': return finite(a + b);
          case '-': return finite(a - b);
          case '*': return finite(a * b);
          case '/': return b === 0 ? error('#DIV/0!') : finite(a / b);
          case '^': return finite(a ** b);
          default: return undefined;
        }
      }
      case 'call': {
        if (!aggregateNames.has(expression.name)) return undefined;
        let count = 0, sum = 0, min = Infinity, max = -Infinity;
        for (const argument of expression.args) {
          const result = inner(argument);
          if (result === undefined) return undefined;
          const values = isRange(result) ? result : [result];
          for (const value of values) {
            if (isError(value) && expression.name !== 'COUNT') return value;
            if (typeof value !== 'number') continue;
            count++;
            sum += value;
            min = Math.min(value, min);
            max = Math.max(value, max);
          }
        }
        switch (expression.name) {
          case 'COUNT': return count;
          case 'MIN': return count ? min : 0;
          case 'MAX': return count ? max : 0;
          case 'AVERAGE': return count ? finite(sum / count) : error('#DIV/0!');
          default: return finite(sum);
        }
      }
    }
  }

  for (const sheet of document.sheets) for (const cell of sheet.cells) {
    if (cell.content.kind !== 'formula' || cell.requiresExternalCalculation) continue;
    const cache = toCache(cellValue(sheet, cell, 0));
    if (cache) updates.push({ target: { sheetId: sheet.id, row: cell.row, column: cell.column }, cache });
  }
  return updates;
}
