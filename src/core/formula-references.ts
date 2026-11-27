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
  readonly endSheet?: string;
  readonly first: Endpoint;
  readonly last: Endpoint;
  readonly range: boolean;
}
export type StructuredSpecifier = '#All' | '#Data' | '#Headers' | '#Totals' | '#This Row' | '@';
export interface StructuredSpan {
  readonly start: number;
  readonly end: number;
}
export interface StructuredReference {
  readonly start: number;
  readonly end: number;
  readonly table?: string;
  readonly specifiers: readonly StructuredSpecifier[];
  readonly startColumn?: string;
  readonly endColumn?: string;
  readonly startColumnSpan?: StructuredSpan;
  readonly endColumnSpan?: StructuredSpan;
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
  | { kind: 'structured'; reference: StructuredReference }
  | { kind: 'spill'; target: FormulaExpression }
  | { kind: 'invoke'; callee: FormulaExpression; args: readonly FormulaExpression[] }
  | { kind: 'union'; parts: readonly FormulaExpression[] }
  | { kind: 'unary'; operator: string; value: FormulaExpression }
  | { kind: 'binary'; operator: string; left: FormulaExpression; right: FormulaExpression }
  | { kind: 'call'; name: string; args: readonly FormulaExpression[] };

const unquotedSheet = '[\\p{L}_][\\p{L}\\p{N}_.]*';
const quotedSheet = "'(?:[^']|'')+'";
const sheet = `(?:${quotedSheet}|${unquotedSheet})`;
const cell = '\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]*';
const row = '\\$?[1-9][0-9]*';
const col = '\\$?[A-Za-z]{1,3}';
const range = `(?:${cell}(?:\\s*:\\s*${cell})?|${row}\\s*:\\s*${row}|${col}\\s*:\\s*${col})`;
const threeDSheetPrefix = `(?:(${sheet})\\s*:\\s*(${sheet})|'((?:[^':]|'')+):((?:[^':]|'')+)')!`;
const threeDReferences = new RegExp(`${threeDSheetPrefix}(${range})(?![\\p{L}\\p{N}_.\\$]|\\s*\\()`, 'uy');
const references = new RegExp(`(?:(${sheet})!)?(${range})(?![\\p{L}\\p{N}_.\\$]|\\s*\\()`, 'uy');
const unsupportedReference = new RegExp(`(?:\\[[^\\]]+\\]${sheet}!?)`, 'uy');
const identifierPattern = /[\p{L}_][\p{L}\p{N}_.]*/uy;
const other = /\s+|"(?:[^"]|"")*"|#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!)|(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?|[\p{L}_][\p{L}\p{N}_.]*|<=|>=|<>|[(){};+*/%^&=<>:#@,-]/iuy;
const precedence: Readonly<Record<string, number>> = {
  '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
  '&': 2,
  '+': 3, '-': 3,
  '*': 4, '/': 4,
  '^': 5,
  ' ': 7,
  ':': 8,
};
const malformed = () => new DocumentError('Clipboard operation cannot transform malformed or unsupported reference syntax.');
export class UnsupportedFormulaReference extends DocumentError {}

const specifierMap: Readonly<Record<string, StructuredSpecifier>> = {
  '#ALL': '#All',
  '#DATA': '#Data',
  '#HEADERS': '#Headers',
  '#TOTALS': '#Totals',
  '#THIS ROW': '#This Row',
  '@': '@',
};

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

function unquoteSheetToken(token: string): string {
  return token.startsWith("'") ? token.slice(1, -1).replaceAll("''", "'") : token;
}

function scanBracketGroup(source: string, start: number): number {
  let depth = 0;
  for (let i = start; i < source.length; i++) {
    const ch = source[i]!;
    if (ch === "'" && i + 1 < source.length) {
      i++;
      continue;
    }
    if (ch === '[') depth++;
    else if (ch === ']') {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  throw malformed();
}

function unescapeStructured(text: string): string {
  let out = '';
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "'" && i + 1 < text.length) out += text[++i];
    else out += text[i];
  }
  return out;
}

export function escapeStructuredColumn(name: string): string {
  return name.replace(/['[\]#@]/gu, match => `'${match}`);
}

function splitTopLevel(source: string, from: number, to: number, delimiter: string): { start: number; end: number }[] {
  const parts: { start: number; end: number }[] = [];
  let depth = 0;
  let partStart = from;
  for (let i = from; i < to; i++) {
    const ch = source[i]!;
    if (ch === "'" && i + 1 < to) {
      i++;
      continue;
    }
    if (ch === '[') depth++;
    else if (ch === ']') depth--;
    else if (depth === 0 && ch === delimiter) {
      parts.push({ start: partStart, end: i });
      partStart = i + 1;
    }
  }
  parts.push({ start: partStart, end: to });
  return parts;
}

function trimSpan(source: string, start: number, end: number): { start: number; end: number } {
  while (start < end && /\s/u.test(source[start]!)) start++;
  while (end > start && /\s/u.test(source[end - 1]!)) end--;
  return { start, end };
}

function unwrapBrackets(source: string, span: { start: number; end: number }): { start: number; end: number } {
  const trimmed = trimSpan(source, span.start, span.end);
  if (trimmed.end - trimmed.start >= 2 && source[trimmed.start] === '[' && source[trimmed.end - 1] === ']') {
    return { start: trimmed.start + 1, end: trimmed.end - 1 };
  }
  return trimmed;
}

function parseStructuredReference(source: string, refStart: number, bracketStart: number, end: number, table?: string): StructuredReference {
  const innerStart = bracketStart + 1;
  const innerEnd = end - 1;
  const specifiers: StructuredSpecifier[] = [];
  let startColumn: string | undefined;
  let endColumn: string | undefined;
  let startColumnSpan: StructuredSpan | undefined;
  let endColumnSpan: StructuredSpan | undefined;

  const trimmedInner = trimSpan(source, innerStart, innerEnd);
  if (trimmedInner.start < trimmedInner.end) {
    if (source[trimmedInner.start] === '@') {
      specifiers.push('@');
      const afterAt = trimSpan(source, trimmedInner.start + 1, trimmedInner.end);
      if (afterAt.start < afterAt.end) {
        const rangeParts = splitTopLevel(source, afterAt.start, afterAt.end, ':');
        const firstSpan = unwrapBrackets(source, rangeParts[0]!);
        startColumn = unescapeStructured(source.slice(firstSpan.start, firstSpan.end));
        startColumnSpan = firstSpan;
        if (rangeParts.length === 2) {
          const secondSpan = unwrapBrackets(source, rangeParts[1]!);
          endColumn = unescapeStructured(source.slice(secondSpan.start, secondSpan.end));
          endColumnSpan = secondSpan;
        } else if (rangeParts.length > 2) throw malformed();
      }
    } else {
      const items = splitTopLevel(source, trimmedInner.start, trimmedInner.end, ',');
      for (const item of items) {
        const cleaned = trimSpan(source, item.start, item.end);
        if (cleaned.start >= cleaned.end) throw malformed();
        const rangeParts = splitTopLevel(source, cleaned.start, cleaned.end, ':');
        if (rangeParts.length === 2) {
          const firstSpan = unwrapBrackets(source, rangeParts[0]!);
          const secondSpan = unwrapBrackets(source, rangeParts[1]!);
          startColumn = unescapeStructured(source.slice(firstSpan.start, firstSpan.end));
          endColumn = unescapeStructured(source.slice(secondSpan.start, secondSpan.end));
          startColumnSpan = firstSpan;
          endColumnSpan = secondSpan;
          continue;
        }
        if (rangeParts.length > 2) throw malformed();
        const span = unwrapBrackets(source, cleaned);
        const raw = unescapeStructured(source.slice(span.start, span.end));
        const spec = specifierMap[raw.toUpperCase()];
        if (spec) specifiers.push(spec);
        else {
          startColumn = raw;
          startColumnSpan = span;
        }
      }
    }
  }

  return {
    start: refStart,
    end,
    ...(table !== undefined && { table }),
    specifiers,
    ...(startColumn !== undefined && { startColumn }),
    ...(endColumn !== undefined && { endColumn }),
    ...(startColumnSpan && { startColumnSpan }),
    ...(endColumnSpan && { endColumnSpan }),
  };
}

interface Token {
  readonly text: string;
  readonly start: number;
  readonly end: number;
  readonly hadSpaceBefore: boolean;
  readonly ref?: FormulaReference;
  readonly structured?: StructuredReference;
}

export interface ParsedFormula {
  readonly expression: FormulaExpression;
  readonly references: readonly FormulaReference[];
  readonly structuredReferences: readonly StructuredReference[];
  readonly nameSpans: readonly { readonly start: number; readonly end: number; readonly name: string }[];
}

function parseFormulaInternal(source: string): ParsedFormula {
  if (!source.startsWith('=') || source.length > documentLimits.formulaLength) throw malformed();
  const refs: FormulaReference[] = [];
  const structuredRefs: StructuredReference[] = [];
  const tokens: Token[] = [];
  let hadSpaceBefore = false;

  for (let start = 1; start < source.length;) {
    unsupportedReference.lastIndex = start;
    if (unsupportedReference.test(source)) {
      throw new UnsupportedFormulaReference('External workbook references are unsupported.');
    }

    threeDReferences.lastIndex = start;
    const threeD = threeDReferences.exec(source);
    if (threeD) {
      const [text, sepFirst, sepLast, combFirst, combLast, body] = threeD;
      const sheetName = sepFirst !== undefined ? unquoteSheetToken(sepFirst) : combFirst!.replaceAll("''", "'");
      const endSheetName = sepLast !== undefined ? unquoteSheetToken(sepLast) : combLast!.replaceAll("''", "'");
      if (/[\[\]:]/u.test(sheetName) || /[\[\]:]/u.test(endSheetName)) throw malformed();
      const [first, last] = body!.split(/\s*:\s*/u);
      const ref: FormulaReference = {
        start, end: threeDReferences.lastIndex,
        first: endpoint(first!), last: endpoint(last ?? first!), range: last !== undefined,
        prefix: text.slice(0, text.length - body!.length),
        sheet: sheetName,
        endSheet: endSheetName,
      };
      refs.push(ref);
      tokens.push({ text, start, end: threeDReferences.lastIndex, hadSpaceBefore, ref });
      hadSpaceBefore = false;
      start = threeDReferences.lastIndex;
      continue;
    }

    references.lastIndex = start;
    const match = references.exec(source);
    if (match) {
      const [text, qualifier, body] = match;
      const [first, last] = body!.split(/\s*:\s*/u);
      const ref: FormulaReference = {
        start, end: references.lastIndex, first: endpoint(first!), last: endpoint(last ?? first!), range: last !== undefined,
        prefix: qualifier ? qualifier + '!' : '',
        ...(qualifier && { sheet: unquoteSheetToken(qualifier) }),
      };
      if (ref.sheet && /[\[\]:]/u.test(ref.sheet)) throw malformed();
      refs.push(ref);
      tokens.push({ text, start, end: references.lastIndex, hadSpaceBefore, ref });
      hadSpaceBefore = false;
      start = references.lastIndex;
      continue;
    }

    identifierPattern.lastIndex = start;
    const idMatch = identifierPattern.exec(source);
    if (idMatch && source[identifierPattern.lastIndex] === '[') {
      const bracketEnd = scanBracketGroup(source, identifierPattern.lastIndex);
      const structured = parseStructuredReference(source, start, identifierPattern.lastIndex, bracketEnd, idMatch[0]);
      structuredRefs.push(structured);
      tokens.push({ text: source.slice(start, bracketEnd), start, end: bracketEnd, hadSpaceBefore, structured });
      hadSpaceBefore = false;
      start = bracketEnd;
      continue;
    }

    if (source[start] === '[') {
      const bracketEnd = scanBracketGroup(source, start);
      const structured = parseStructuredReference(source, start, start, bracketEnd);
      structuredRefs.push(structured);
      tokens.push({ text: source.slice(start, bracketEnd), start, end: bracketEnd, hadSpaceBefore, structured });
      hadSpaceBefore = false;
      start = bracketEnd;
      continue;
    }

    other.lastIndex = start;
    const rest = other.exec(source);
    if (!rest) throw malformed();
    if (/^\s/u.test(rest[0])) {
      hadSpaceBefore = true;
    } else {
      tokens.push({ text: rest[0], start, end: other.lastIndex, hadSpaceBefore });
      hadSpaceBefore = false;
    }
    start = other.lastIndex;
  }

  const nameSpans: { start: number; end: number; name: string }[] = [];
  let index = 0;
  const peekToken = () => tokens[index];
  const peek = () => tokens[index]?.text ?? '';
  function take(expected?: string) {
    const token = tokens[index++];
    if (!token || expected !== undefined && token.text !== expected) throw malformed();
    return token;
  }
  function constant(text: string): FormulaConstant | undefined {
    if (text.startsWith('"')) return { kind: 'literal', value: text.slice(1, -1).replaceAll('""', '"') };
    if (text.startsWith('#') && text.length > 1) return { kind: 'error', value: text.toUpperCase() };
    if (/^[0-9.]/u.test(text)) return { kind: 'literal', value: Number(text) };
    if (/^(?:true|false)$/iu.test(text)) return { kind: 'literal', value: text.toUpperCase() === 'TRUE' };
  }
  function array(): FormulaExpression {
    const rows: FormulaConstant[][] = [[]];
    for (;;) {
      const token = take();
      if (token.ref || token.structured) throw malformed();
      let value: FormulaConstant | undefined;
      if (token.text === '+' || token.text === '-') {
        const next = take();
        const numeric = next.ref || next.structured ? undefined : constant(next.text);
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
  function parseCallArgs(depth: number): FormulaExpression[] {
    take('(');
    const args: FormulaExpression[] = [];
    if (peek() !== ')') {
      do {
        if (args.length) take(',');
        args.push(peek() === ',' || peek() === ')' ? { kind: 'missing' } : expression(1, depth + 1));
      } while (peek() === ',');
    }
    take(')');
    return args;
  }
  function primary(depth: number): FormulaExpression {
    if (depth > 64) throw malformed();
    const token = take();
    const { text, ref, structured } = token;
    if (ref) return { kind: 'reference', reference: ref };
    if (structured) return { kind: 'structured', reference: structured };
    if (text === '{') return array();
    if (text === '(') {
      const parts = [expression(1, depth + 1)];
      while (peek() === ',') { take(','); parts.push(expression(1, depth + 1)); }
      take(')');
      return parts.length === 1 ? parts[0]! : { kind: 'union', parts };
    }
    if (text === '+' || text === '-') return { kind: 'unary', operator: text, value: postfix(depth + 1) };
    if (text === '@') return { kind: 'unary', operator: '@', value: postfix(depth + 1) };
    const value = constant(text);
    if (value) return value;
    if (!/^[\p{L}_]/u.test(text)) throw malformed();
    if (peek() !== '(') {
      nameSpans.push({ start: token.start, end: token.end, name: text });
      return { kind: 'name', name: text };
    }
    const args = parseCallArgs(depth);
    return { kind: 'call', name: text.toUpperCase(), args };
  }
  function postfix(depth: number): FormulaExpression {
    let left = primary(depth);
    for (;;) {
      const next = peek();
      if (next === '#') {
        take('#');
        left = { kind: 'spill', target: left };
      } else if (next === '%') {
        take('%');
        left = { kind: 'unary', operator: '%', value: left };
      } else if (next === '(' && (left.kind === 'call' || left.kind === 'invoke')) {
        left = { kind: 'invoke', callee: left, args: parseCallArgs(depth) };
      } else break;
    }
    return left;
  }
  function isReferenceCapable(node: FormulaExpression): boolean {
    switch (node.kind) {
      case 'reference':
      case 'structured':
      case 'name':
      case 'call':
      case 'invoke':
      case 'union':
      case 'spill':
        return true;
      case 'binary':
        return node.operator === ':' || node.operator === ' ';
      default:
        return false;
    }
  }
  function canStartReference(token: Token | undefined): boolean {
    if (!token) return false;
    if (token.ref || token.structured || token.text === '(') return true;
    return /^[\p{L}_]/u.test(token.text) && !/^(?:true|false)$/iu.test(token.text);
  }
  function expression(min: number, depth: number): FormulaExpression {
    let left = postfix(depth);
    for (;;) {
      const next = peek();
      const opPrecedence = precedence[next] ?? 0;
      if (opPrecedence >= min && (next !== ':' || isReferenceCapable(left))) {
        const operator = take().text;
        const right = expression(opPrecedence + 1, depth + 1);
        left = { kind: 'binary', operator, left, right };
        continue;
      }
      const nextToken = peekToken();
      if (min <= 7 && nextToken?.hadSpaceBefore && isReferenceCapable(left) && canStartReference(nextToken)) {
        const right = expression(8, depth + 1);
        left = { kind: 'binary', operator: ' ', left, right };
        continue;
      }
      break;
    }
    return left;
  }
  const result = expression(1, 0);
  if (index !== tokens.length) throw malformed();
  return { expression: result, references: refs, structuredReferences: structuredRefs, nameSpans };
}

export function parseFormula(source: string): {
  expression: FormulaExpression;
  references: readonly FormulaReference[];
  structuredReferences: readonly StructuredReference[];
} {
  const { expression, references, structuredReferences } = parseFormulaInternal(source);
  return { expression, references, structuredReferences };
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

export function quote3DSheet(start: string, end: string): string {
  const simple = /^[\p{L}_][\p{L}\p{N}_.]*$/u;
  if (simple.test(start) && simple.test(end)) return `${start}:${end}!`;
  return `'${start.replaceAll("'", "''")}:${end.replaceAll("'", "''")}'!`;
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

/** Shift references whose rows lie entirely inside [first, last]. A reference that extends outside that block stays put. */
export function shiftEnclosedRows(source: string, first: number, last: number, dr: number): string {
  return rewrite(source, ref => {
    const rows = [ref.first.row, ref.last.row];
    if (rows.some(row => row === undefined || row < first || row > last)) return source.slice(ref.start, ref.end);
    const shift = (point: Endpoint) => endpointText(point, point.row! + dr, point.column);
    const start = shift(ref.first);
    const end = shift(ref.last);
    return start === '#REF!' || end === '#REF!' ? '#REF!' : ref.prefix + start + (ref.range ? ':' + end : '');
  });
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
    if (ref.endSheet !== undefined) return source.slice(ref.start, ref.end);
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

export function rewriteSheetReferences(
  source: string,
  sheet: string,
  replacement?: string,
  sheetOrder?: readonly string[],
): string {
  return rewrite(source, ref => {
    if (ref.endSheet !== undefined) {
      const startMatch = ref.sheet !== undefined && sameSheet(ref.sheet, sheet);
      const endMatch = sameSheet(ref.endSheet, sheet);
      if (!startMatch && !endMatch) return source.slice(ref.start, ref.end);
      const localRef = source.slice(ref.start + ref.prefix.length, ref.end);
      if (replacement !== undefined) {
        const nextStart = startMatch ? replacement : ref.sheet!;
        const nextEnd = endMatch ? replacement : ref.endSheet;
        return quote3DSheet(nextStart, nextEnd) + localRef;
      }
      if (!sheetOrder || sheetOrder.length === 0) return '#REF!';
      const startIdx = sheetOrder.findIndex(name => sameSheet(name, ref.sheet!));
      const endIdx = sheetOrder.findIndex(name => sameSheet(name, ref.endSheet!));
      if (startIdx < 0 || endIdx < 0) return '#REF!';
      const lo = Math.min(startIdx, endIdx);
      const hi = Math.max(startIdx, endIdx);
      const remaining = sheetOrder.slice(lo, hi + 1).filter(name => !sameSheet(name, sheet));
      if (remaining.length === 0) return '#REF!';
      if (remaining.length === 1) return quoteSheet(remaining[0]!) + localRef;
      const forward = startIdx <= endIdx;
      return quote3DSheet(forward ? remaining[0]! : remaining.at(-1)!, forward ? remaining.at(-1)! : remaining[0]!) + localRef;
    }
    if (ref.sheet === undefined || !sameSheet(ref.sheet, sheet)) return source.slice(ref.start, ref.end);
    return replacement === undefined ? '#REF!' :
      quoteSheet(replacement) + source.slice(ref.start + ref.prefix.length, ref.end);
  });
}

function sheetIn3DSpan(ref: FormulaReference, sheet: string, sheetOrder?: readonly string[]): boolean {
  if (ref.endSheet === undefined) return false;
  if (ref.sheet !== undefined && sameSheet(ref.sheet, sheet) || sameSheet(ref.endSheet, sheet)) return true;
  if (!sheetOrder || ref.sheet === undefined) return false;
  const startIdx = sheetOrder.findIndex(name => sameSheet(name, ref.sheet!));
  const endIdx = sheetOrder.findIndex(name => sameSheet(name, ref.endSheet!));
  const targetIdx = sheetOrder.findIndex(name => sameSheet(name, sheet));
  return startIdx >= 0 && endIdx >= 0 && targetIdx >= Math.min(startIdx, endIdx) && targetIdx <= Math.max(startIdx, endIdx);
}

export function shiftReferences(
  source: string,
  localSheet: string,
  sheet: string,
  edit: AxisEdit,
  sheetOrder?: readonly string[],
): string {
  return rewrite(source, ref => {
    const original = source.slice(ref.start, ref.end);
    const applies = ref.endSheet !== undefined
      ? sheetIn3DSpan(ref, sheet, sheetOrder)
      : sameSheet(ref.sheet ?? localSheet, sheet);
    if (!applies || ref.first[edit.axis] === undefined) return original;
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

function applySpanReplacements(source: string, replacements: readonly { start: number; end: number; text: string }[]): string {
  if (replacements.length === 0) return source;
  const ordered = [...replacements].sort((a, b) => a.start - b.start);
  const pieces: string[] = [];
  let cursor = 0;
  for (const edit of ordered) {
    pieces.push(source.slice(cursor, edit.start), edit.text);
    cursor = edit.end;
  }
  pieces.push(source.slice(cursor));
  return pieces.join('');
}

export function renameTableReferences(source: string, oldName: string, newName: string): string {
  let parsed: ParsedFormula;
  try {
    parsed = parseFormulaInternal(source);
  } catch {
    return source;
  }
  const key = ordinalNameKey(oldName);
  const edits: { start: number; end: number; text: string }[] = [];
  for (const ref of parsed.structuredReferences) {
    if (ref.table && ordinalNameKey(ref.table) === key) {
      edits.push({ start: ref.start, end: ref.start + ref.table.length, text: newName });
    }
  }
  for (const span of parsed.nameSpans) {
    if (ordinalNameKey(span.name) === key) {
      edits.push({ start: span.start, end: span.end, text: newName });
    }
  }
  return applySpanReplacements(source, edits);
}

export function renameTableColumnReferences(
  source: string,
  tableName: string,
  oldColumn: string,
  newColumn: string,
  formulaLocation?: { readonly sheetId: string; readonly row: number; readonly column: number },
  tableArea?: { readonly sheetId: string; readonly row: number; readonly column: number; readonly rows: number; readonly columns: number },
): string {
  let parsed: ParsedFormula;
  try {
    parsed = parseFormulaInternal(source);
  } catch {
    return source;
  }
  const tableKey = ordinalNameKey(tableName);
  const colKey = ordinalNameKey(oldColumn);
  const escaped = escapeStructuredColumn(newColumn);
  const edits: { start: number; end: number; text: string }[] = [];
  const insideTable = Boolean(
    formulaLocation && tableArea && formulaLocation.sheetId === tableArea.sheetId &&
    formulaLocation.row >= tableArea.row && formulaLocation.row < tableArea.row + tableArea.rows &&
    formulaLocation.column >= tableArea.column && formulaLocation.column < tableArea.column + tableArea.columns,
  );
  for (const ref of parsed.structuredReferences) {
    const matchesTable = ref.table ? ordinalNameKey(ref.table) === tableKey : insideTable;
    if (!matchesTable) continue;
    if (ref.startColumn !== undefined && ref.startColumnSpan && ordinalNameKey(ref.startColumn) === colKey) {
      edits.push({ start: ref.startColumnSpan.start, end: ref.startColumnSpan.end, text: escaped });
    }
    if (ref.endColumn !== undefined && ref.endColumnSpan && ordinalNameKey(ref.endColumn) === colKey) {
      edits.push({ start: ref.endColumnSpan.start, end: ref.endColumnSpan.end, text: escaped });
    }
  }
  return applySpanReplacements(source, edits);
}
