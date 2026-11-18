import { DocumentError, documentLimits, type Coordinate } from './workbook';
import { ordinalNameKey } from './workbook-validation';
import { shiftBox, type AxisEdit } from './axis-edit';

interface Endpoint {
  readonly row?: number;
  readonly column?: number;
  readonly lockRow: boolean;
  readonly lockColumn: boolean;
}
export interface FormulaReference {
  readonly start: number;
  readonly end: number;
  readonly prefix: string;
  readonly sheet?: string;
  readonly first: Endpoint;
  readonly last: Endpoint;
  readonly range: boolean;
}
export interface ReferenceBox {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

type FormulaConstant = { kind: 'literal'; value: string | number | boolean } | { kind: 'error'; value: string };

export type FormulaExpression =
  | FormulaConstant
  | { kind: 'missing' }
  | { kind: 'array'; rows: readonly (readonly FormulaConstant[])[] }
  | { kind: 'name'; name: string }
  | { kind: 'reference'; reference: FormulaReference }
  | { kind: 'union'; parts: readonly FormulaExpression[] }
  | { kind: 'unary'; operator: string; value: FormulaExpression }
  | { kind: 'binary'; operator: string; left: FormulaExpression; right: FormulaExpression }
  | { kind: 'call'; name: string; args: readonly FormulaExpression[] };

const sheet = "(?:'(?:[^']|'')+'|[\\p{L}_][\\p{L}\\p{N}_.]*)";
const cell = '\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]*';
const row = '\\$?[1-9][0-9]*';
const col = '\\$?[A-Za-z]{1,3}';
const range = `(?:${cell}(?:\\s*:\\s*${cell})?|${row}\\s*:\\s*${row}|${col}\\s*:\\s*${col})`;
const references = new RegExp(`(?:(${sheet})!)?(${range})(?![\\p{L}\\p{N}_.\\$]|\\s*\\()`, 'uy');
const unsupportedReference = new RegExp(`(?:\\[\\d+\\]${sheet}!?|(?:${sheet}:${sheet}|'[^']+:[^']+')!)`, 'uy');
const other = /\s+|"(?:[^"]|"")*"|#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)|(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?|[\p{L}_][\p{L}\p{N}_.]*(?:\[[^\[\]]+\])?|<=|>=|<>|[(){};+*/%^&=<>,-]/iuy;
const precedence: Readonly<Record<string, number>> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
const malformed = () => new DocumentError('Clipboard operation cannot transform malformed or unsupported reference syntax.');
export class UnsupportedFormulaReference extends DocumentError {}

export function columnLetters(column: number): string {
  let label = '';
  for (let n = column; n > 0; n = Math.floor((n - 1) / 26)) label = String.fromCharCode(65 + (n - 1) % 26) + label;
  return label;
}

export function columnIndex(letters: string): number {
  return [...letters.toUpperCase()].reduce((n, letter) => n * 26 + letter.charCodeAt(0) - 64, 0);
}

function endpoint(text: string): Endpoint {
  const match = /^(\$?)([A-Za-z]*)(\$?)(\d*)$/u.exec(text)!;
  const [, firstLock, letters, secondLock, digits] = match;
  return {
    ...(letters && { column: columnIndex(letters) }), ...(digits && { row: Number(digits) }),
    lockColumn: !!letters && firstLock === '$', lockRow: (letters ? secondLock : firstLock) === '$',
  };
}

export function parseFormula(source: string): { expression: FormulaExpression; references: readonly FormulaReference[] } {
  if (!source.startsWith('=') || source.length > documentLimits.formulaLength) throw malformed();
  const refs: FormulaReference[] = [];
  const tokens: { text: string; ref?: FormulaReference }[] = [];
  for (let start = 1; start < source.length;) {
    unsupportedReference.lastIndex = start;
    if (unsupportedReference.test(source)) {
      throw new UnsupportedFormulaReference('External workbook and 3D references are unsupported.');
    }
    references.lastIndex = start;
    const match = references.exec(source);
    if (match) {
      const [text, qualifier, body] = match;
      const [first, last] = body!.split(/\s*:\s*/u);
      const ref: FormulaReference = {
        start, end: references.lastIndex, first: endpoint(first!), last: endpoint(last ?? first!), range: last !== undefined,
        prefix: qualifier ? qualifier + '!' : '',
        ...(qualifier && { sheet: qualifier.startsWith("'") ? qualifier.slice(1, -1).replaceAll("''", "'") : qualifier }),
      };
      if (ref.sheet && /[\[\]:]/u.test(ref.sheet)) throw malformed();
      refs.push(ref);
      tokens.push({ text, ref });
      start = references.lastIndex;
      continue;
    }
    other.lastIndex = start;
    const rest = other.exec(source);
    if (!rest) throw malformed();
    if (!/^\s/u.test(rest[0])) tokens.push({ text: rest[0] });
    start = other.lastIndex;
  }
  let index = 0;
  const peek = () => tokens[index]?.text ?? '';
  function take(expected?: string) {
    const token = tokens[index++];
    if (!token || expected !== undefined && token.text !== expected) throw malformed();
    return token;
  }
  function constant(text: string): FormulaConstant | undefined {
    if (text.startsWith('"')) return { kind: 'literal', value: text.slice(1, -1).replaceAll('""', '"') };
    if (text.startsWith('#')) return { kind: 'error', value: text.toUpperCase() };
    if (/^[0-9.]/u.test(text)) return { kind: 'literal', value: Number(text) };
    if (/^(?:true|false)$/iu.test(text)) return { kind: 'literal', value: text.toUpperCase() === 'TRUE' };
  }
  function array(): FormulaExpression {
    const rows: FormulaConstant[][] = [[]];
    for (;;) {
      const token = take();
      if (token.ref) throw malformed();
      let value: FormulaConstant | undefined;
      if (token.text === '+' || token.text === '-') {
        const next = take();
        const numeric = next.ref ? undefined : constant(next.text);
        if (numeric?.kind !== 'literal' || typeof numeric.value !== 'number') throw malformed();
        value = { kind: 'literal', value: token.text === '-' ? -numeric.value : numeric.value };
      } else value = constant(token.text);
      if (!value) throw malformed();
      rows.at(-1)!.push(value);
      const delimiter = take().text;
      if (delimiter === '}') break;
      if (delimiter === ';') rows.push([]);
      else if (delimiter !== ',') throw malformed();
    }
    if (rows.some(row => row.length !== rows[0]!.length)) throw malformed();
    return { kind: 'array', rows };
  }
  function primary(depth: number): FormulaExpression {
    if (depth > 64) throw malformed();
    const { text, ref } = take();
    if (ref) return { kind: 'reference', reference: ref };
    if (text === '{') return array();
    if (text === '(') {
      const parts = [expression(1, depth + 1)];
      while (peek() === ',') { take(','); parts.push(expression(1, depth + 1)); }
      take(')');
      return parts.length === 1 ? parts[0]! : { kind: 'union', parts };
    }
    if (text === '+' || text === '-') return { kind: 'unary', operator: text, value: primary(depth + 1) };
    const value = constant(text);
    if (value) return value;
    if (!/^[\p{L}_]/u.test(text)) throw malformed();
    if (peek() !== '(') return { kind: 'name', name: text };
    take('(');
    const args: FormulaExpression[] = [];
    if (peek() !== ')') {
      do {
        if (args.length) take(',');
        args.push(peek() === ',' || peek() === ')' ? { kind: 'missing' } : expression(1, depth + 1));
      } while (peek() === ',');
    }
    take(')');
    return { kind: 'call', name: text.toUpperCase(), args };
  }
  function expression(min: number, depth: number): FormulaExpression {
    let left = primary(depth);
    while (peek() === '%') {
      take('%');
      left = { kind: 'unary', operator: '%', value: left };
    }
    while ((precedence[peek()] ?? 0) >= min) {
      const operator = take().text;
      const right = expression(precedence[operator]! + 1, depth + 1);
      left = { kind: 'binary', operator, left, right };
    }
    return left;
  }
  const result = expression(1, 0);
  if (index !== tokens.length) throw malformed();
  return { expression: result, references: refs };
}

export function referenceBox({ first, last }: FormulaReference): ReferenceBox {
  return {
    top: first.row === undefined ? 1 : Math.min(first.row, last.row!),
    bottom: first.row === undefined ? documentLimits.rows : Math.max(first.row, last.row!),
    left: first.column === undefined ? 1 : Math.min(first.column, last.column!),
    right: first.column === undefined ? documentLimits.columns : Math.max(first.column, last.column!),
  };
}

export function quoteSheet(name: string): string {
  return "'" + name.replaceAll("'", "''") + "'!";
}

function endpointText(point: Endpoint, row = point.row, column = point.column): string {
  if (row !== undefined && (!Number.isSafeInteger(row) || row < 1 || row > documentLimits.rows) ||
      column !== undefined && (!Number.isSafeInteger(column) || column < 1 || column > documentLimits.columns)) return '#REF!';
  return (column === undefined ? '' : (point.lockColumn ? '$' : '') + columnLetters(column)) +
    (row === undefined ? '' : (point.lockRow ? '$' : '') + row);
}

function rewrite(source: string, visit: (ref: FormulaReference) => string): string {
  const pieces: string[] = [];
  let end = 0;
  for (const ref of parseFormula(source).references) {
    pieces.push(source.slice(end, ref.start), visit(ref));
    end = ref.end;
  }
  pieces.push(source.slice(end));
  return pieces.join('');
}

export function copyReferences(source: string, dr: number, dc: number): string {
  return rewrite(source, ref => {
    const shift = (point: Endpoint) => endpointText(point,
      point.row === undefined ? undefined : point.row + (point.lockRow ? 0 : dr),
      point.column === undefined ? undefined : point.column + (point.lockColumn ? 0 : dc));
    const first = shift(ref.first);
    const last = shift(ref.last);
    return first === '#REF!' || last === '#REF!' ? '#REF!' : ref.prefix + first + (ref.range ? ':' + last : '');
  });
}

export interface ReferenceMove {
  readonly sourceSheet: string;
  readonly destinationSheet: string;
  readonly from: ReferenceBox;
  readonly dr: number;
  readonly dc: number;
}

export function intersection(a: ReferenceBox, b: ReferenceBox): ReferenceBox | undefined {
  const top = Math.max(a.top, b.top), left = Math.max(a.left, b.left);
  const bottom = Math.min(a.bottom, b.bottom), right = Math.min(a.right, b.right);
  return top <= bottom && left <= right ? { top, left, bottom, right } : undefined;
}

export function sameSheet(left: string, right: string): boolean {
  return ordinalNameKey(left) === ordinalNameKey(right);
}

export function moveReferences(source: string, formulaSheet: string, resultSheet: string, move: ReferenceMove): string {
  return rewrite(source, ref => {
    const resolved = ref.sheet ?? formulaSheet;
    const localText = source.slice(ref.start + ref.prefix.length, ref.end);
    const prefix = ref.prefix || (sameSheet(resolved, resultSheet) ? '' : quoteSheet(resolved));
    const original = referenceBox(ref);
    const hit = sameSheet(resolved, move.sourceSheet) ? intersection(original, move.from) : undefined;
    if (!hit) return prefix + localText;
    const { top, bottom, left, right } = original;
    const parts: string[] = [];
    function box(t: number, l: number, b: number, r: number, qualifier: string, dr = 0, dc = 0) {
      const first = endpointText(ref.first, t + dr, l + dc);
      const last = endpointText(ref.last, b + dr, r + dc);
      parts.push(qualifier + first + (t === b && l === r ? '' : ':' + last));
    }
    if (top < hit.top) box(top, left, hit.top - 1, right, prefix);
    if (hit.bottom < bottom) box(hit.bottom + 1, left, bottom, right, prefix);
    if (left < hit.left) box(hit.top, left, hit.bottom, hit.left - 1, prefix);
    if (hit.right < right) box(hit.top, hit.right + 1, hit.bottom, right, prefix);
    const target = ref.sheet || !sameSheet(move.destinationSheet, resultSheet) ? quoteSheet(move.destinationSheet) : '';
    box(hit.top, hit.left, hit.bottom, hit.right, target, move.dr, move.dc);
    return parts.length === 1 ? parts[0]! : '(' + parts.join(',') + ')';
  });
}

export function formulaAddress(point: Coordinate): string {
  return columnLetters(point.column) + point.row;
}

export function rewriteSheetReferences(source: string, sheet: string, replacement?: string): string {
  return rewrite(source, ref => {
    if (ref.sheet === undefined || !sameSheet(ref.sheet, sheet)) return source.slice(ref.start, ref.end);
    return replacement === undefined ? '#REF!' :
      quoteSheet(replacement) + source.slice(ref.start + ref.prefix.length, ref.end);
  });
}

export function shiftReferences(source: string, localSheet: string, sheet: string, edit: AxisEdit): string {
  return rewrite(source, ref => {
    const original = source.slice(ref.start, ref.end);
    if (!sameSheet(ref.sheet ?? localSheet, sheet) || ref.first[edit.axis] === undefined) return original;
    const area = referenceBox(ref);
    const shifted = shiftBox(edit, area);
    if (!shifted?.length) return '#REF!';
    if (shifted[0] === area) return original;
    const parts = shifted.map(part => {
      const reverseRows = (ref.first.row ?? 0) > (ref.last.row ?? 0);
      const reverseColumns = (ref.first.column ?? 0) > (ref.last.column ?? 0);
      const compact = !edit.lane;
      const first = endpointText(ref.first,
        compact && ref.first.row === undefined ? undefined : reverseRows ? part.bottom : part.top,
        compact && ref.first.column === undefined ? undefined : reverseColumns ? part.right : part.left);
      const last = endpointText(ref.last,
        compact && ref.last.row === undefined ? undefined : reverseRows ? part.top : part.bottom,
        compact && ref.last.column === undefined ? undefined : reverseColumns ? part.left : part.right);
      if (first === '#REF!' || last === '#REF!') return '#REF!';
      const wholeAxis = compact && (ref.first.row === undefined || ref.first.column === undefined);
      return ref.prefix + first + (first !== last || wholeAxis ? ':' + last : '');
    });
    return parts.length === 1 ? parts[0]! : '(' + parts.join(',') + ')';
  });
}
