import type { FormulaExpression } from '../core/formula-references';
import { documentLimits, type CellCache } from '../core/workbook';

export type FormulaError = Extract<CellCache, { kind: 'error' }>;
export type Scalar = number | string | boolean | null | FormulaError;
export interface Range {
  readonly sheetId: string;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}
export interface FormulaArray {
  readonly kind: 'array';
  readonly rows: number;
  readonly columns: number;
  readonly at: (row: number, column: number) => Scalar;
}
export interface FormulaOmitted {
  readonly kind: 'omitted';
}
export interface FormulaLambda {
  readonly kind: 'lambda';
  readonly params: readonly string[];
  readonly body: FormulaExpression;
  readonly env: ReadonlyMap<string, FormulaValue>;
}
export type FormulaValue = Scalar | readonly Range[] | FormulaArray | FormulaOmitted | FormulaLambda;

export const omitted: FormulaOmitted = Object.freeze({ kind: 'omitted' });
export const error = (value: FormulaError['value']): FormulaError => ({ kind: 'error', value });
export const isError = (value: unknown): value is FormulaError =>
  typeof value === 'object' && value !== null && !Array.isArray(value) && 'kind' in value && value.kind === 'error';
export const isRange = (value: FormulaValue): value is readonly Range[] => Array.isArray(value);
export const isArray = (value: FormulaValue): value is FormulaArray =>
  typeof value === 'object' && value !== null && !Array.isArray(value) && 'kind' in value && value.kind === 'array';
export const isOmitted = (value: FormulaValue): value is FormulaOmitted =>
  typeof value === 'object' && value !== null && !Array.isArray(value) && 'kind' in value && value.kind === 'omitted';
export const isLambda = (value: FormulaValue): value is FormulaLambda =>
  typeof value === 'object' && value !== null && !Array.isArray(value) && 'kind' in value && value.kind === 'lambda';
export const finite = (value: number): Scalar => Number.isFinite(value) ? value || 0 : error('#NUM!');

export function arrayFromMatrix(rows: number, columns: number, cells: readonly Scalar[]): FormulaArray {
  return {
    kind: 'array',
    rows,
    columns,
    at: (row, column) => cells[row * columns + column] ?? null,
  };
}

const numericText = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/iu;

export function numberValue(value: Scalar, decimal = '.', group = ''): number {
  if (value === null || typeof value === 'boolean' || typeof value === 'number') return Number(value);
  if (typeof value !== 'string') return NaN;
  const localized = value.trim();
  const normalized = group === '\u00a0' || group === '\u202f' ? localized.replaceAll(' ', group) : localized;
  const boundary = normalized.search(/[eE]/u);
  const mantissa = boundary < 0 ? normalized : normalized.slice(0, boundary);
  const exponent = boundary < 0 ? '' : normalized.slice(boundary);
  const [integer, fraction, ...extra] = mantissa.split(decimal);
  if (group && integer!.includes(group) && !/^[+-]?\d/u.test(integer!)) return NaN;
  const ungrouped = group ? integer!.replaceAll(group, '') : integer!;
  if (extra.length || decimal !== '.' && (ungrouped.includes('.') || fraction?.includes('.'))) return NaN;
  const canonical = ungrouped + (fraction === undefined ? '' : '.' + fraction) + exponent;
  return numericText.test(canonical) ? Number(canonical) : NaN;
}
export function textValue(value: Exclude<Scalar, FormulaError>, decimal = '.'): string {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') return String(value).replace('.', decimal);
  return String(value ?? '');
}
export function scalarCache(value: Scalar): CellCache {
  if (isError(value)) return value;
  if (value === null) return { kind: 'number', value: 0 };
  if (typeof value === 'number') return { kind: 'number', value };
  if (typeof value === 'boolean') return { kind: 'boolean', value };
  return value.length <= documentLimits.textLength && !value.includes('\0') ? { kind: 'text', value } : error('#VALUE!');
}

export function* arrayValues(array: FormulaArray): Iterable<Scalar> {
  for (let row = 0; row < array.rows; row++) for (let column = 0; column < array.columns; column++) {
    yield array.at(row, column);
  }
}

export function compare(a: Exclude<Scalar, FormulaError>, b: Exclude<Scalar, FormulaError>): number {
  const left = a ?? (typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0);
  const right = b ?? (typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0);
  const rank = (value: typeof left) => typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2;
  if (typeof left !== typeof right) return rank(left) - rank(right);
  const l = typeof left === 'string' ? left.toUpperCase() : left;
  const r = typeof right === 'string' ? right.toUpperCase() : right;
  return l === r ? 0 : l < r ? -1 : 1;
}
export function matchesComparison(order: number, operator: string): boolean {
  switch (operator) {
    case '=': return order === 0;
    case '<>': return order !== 0;
    case '>': return order > 0;
    case '<': return order < 0;
    case '>=': return order >= 0;
    case '<=': return order <= 0;
    default: return false;
  }
}

export function* formulaNodes(expression: FormulaExpression): Iterable<FormulaExpression> {
  const remaining = [expression];
  while (remaining.length) {
    const node = remaining.pop()!;
    yield node;
    switch (node.kind) {
      case 'call': remaining.push(...node.args); break;
      case 'invoke': remaining.push(...node.args, node.callee); break;
      case 'spill': remaining.push(node.target); break;
      case 'union': remaining.push(...node.parts); break;
      case 'binary': remaining.push(node.right, node.left); break;
      case 'unary': remaining.push(node.value); break;
    }
  }
}
