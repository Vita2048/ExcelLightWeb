/** Feasibility subset. S07–S09 extend the grammar, transformations and evaluator. */
export type FormulaValue = number | string | boolean | null | { error: string };
type Reference = { text: string; start: number; end: number };
type Expression =
  | { kind: 'literal'; value: FormulaValue }
  | { kind: 'reference'; first: Reference; last?: Reference }
  | { kind: 'binary'; operator: string; left: Expression; right: Expression }
  | { kind: 'sign'; operator: string; value: Expression }
  | { kind: 'call'; name: string; args: Expression[] };

const addressPattern = /^(?:(?:'(?:[^']|'')+'|[\w.]+)!)?(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/u;
const tokenPattern = /\s+|"(?:""|[^"])*"|(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!\$?[A-Za-z]{1,3}\$?[1-9]\d*|#[A-Z0-9/?!]+|\$?[A-Za-z]{1,3}\$[1-9]\d*|\$[A-Za-z]{1,3}[1-9]\d*|[A-Za-z_][\w.]*|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|[()+*/,:-]/uy;
const priorities: Record<string, number> = { '+': 1, '-': 1, '*': 2, '/': 2 };
const error = (code: string): FormulaValue => ({ error: code });
const isError = (value: FormulaValue): value is { error: string } => value !== null && typeof value === 'object';

function parse(source: string): { expression: Expression; references: Reference[] } {
  if (!source.startsWith('=') || source.length > 8192) throw new Error('Invalid formula source');
  const pattern = new RegExp(tokenPattern);
  const references: Reference[] = [];
  const tokens: Reference[] = [];
  for (let offset = 1; offset < source.length;) {
    pattern.lastIndex = offset;
    const match = pattern.exec(source);
    if (!match) throw new Error('Invalid formula token');
    if (!/^\s/u.test(match[0])) tokens.push({ text: match[0], start: offset, end: pattern.lastIndex });
    offset = pattern.lastIndex;
  }
  let cursor = 0;
  const peek = () => tokens[cursor]?.text ?? '';
  function take(expected?: string) {
    const token = tokens[cursor++];
    if (!token || (expected !== undefined && token.text !== expected)) throw new Error('Unexpected formula token');
    return token;
  }
  function primary(depth: number): Expression {
    if (depth > 64) throw new Error('Formula nesting limit');
    const token = take();
    const text = token.text;
    if (text === '(') {
      const value = expression(1, depth + 1);
      take(')');
      return value;
    }
    if (text === '+' || text === '-') return { kind: 'sign', operator: text, value: primary(depth + 1) };
    if (text.startsWith('"')) return { kind: 'literal', value: text.slice(1, -1).replaceAll('""', '"') };
    if (text.startsWith('#')) return { kind: 'literal', value: error(text) };
    if (/^(true|false)$/iu.test(text)) return { kind: 'literal', value: text.toUpperCase() === 'TRUE' };
    if (/^[\d.]/u.test(text)) {
      const value = Number(text);
      return { kind: 'literal', value: Number.isFinite(value) ? value : error('#NUM!') };
    }
    if (peek() === '(' && /^[A-Za-z_][\w.]*$/u.test(text)) {
      take('(');
      const args: Expression[] = [];
      if (peek() !== ')') {
        do {
          if (args.length) take(',');
          args.push(expression(1, depth + 1));
        } while (peek() === ',');
      }
      take(')');
      return { kind: 'call', name: text.toUpperCase(), args };
    }
    if (!addressPattern.test(text)) throw new Error('Expected a cell reference');
    references.push(token);
    if (peek() === ':') {
      take(':');
      const last = take();
      if (!addressPattern.test(last.text)) throw new Error('Expected end of range');
      references.push(last);
      return { kind: 'reference', first: token, last };
    }
    return { kind: 'reference', first: token };
  }
  function expression(min: number, depth: number): Expression {
    let left = primary(depth);
    while ((priorities[peek()] ?? 0) >= min) {
      const operator = take().text;
      const right = expression(priorities[operator]! + 1, depth + 1);
      left = { kind: 'binary', operator, left, right };
    }
    return left;
  }
  const result = expression(1, 0);
  if (cursor !== tokens.length) throw new Error('Unexpected formula suffix');
  return { expression: result, references };
}

function coordinates(text: string) {
  const match = addressPattern.exec(text);
  if (!match) throw new Error('Invalid reference');
  const [, colLock, letters, rowLock, digits] = match;
  const col = [...letters!.toUpperCase()].reduce((index, letter) => index * 26 + letter.charCodeAt(0) - 64, 0);
  return { prefix: text.slice(0, text.lastIndexOf('!') + 1), colLock, col, rowLock, row: Number(digits) };
}

function letters(index: number): string {
  let result = '';
  for (let col = index; col > 0; col = Math.floor((col - 1) / 26)) {
    result = String.fromCharCode(65 + (col - 1) % 26) + result;
  }
  return result;
}

function transform(source: string, map: (ref: ReturnType<typeof coordinates>) => { row: number; col: number }) {
  let updated = source;
  for (const token of parse(source).references.reverse()) {
    const ref = coordinates(token.text);
    const { row, col } = map(ref);
    const text = row >= 1 && row <= 10_000 && col >= 1 && col <= 100
      ? `${ref.prefix}${ref.colLock}${letters(col)}${ref.rowLock}${row}`
      : '#REF!';
    updated = updated.slice(0, token.start) + text + updated.slice(token.end);
  }
  return updated;
}

export function copyFormula(source: string, origin: string, destination: string): string {
  const a = coordinates(origin);
  const b = coordinates(destination);
  return transform(source, ref => ({
    row: ref.row + (ref.rowLock ? 0 : b.row - a.row),
    col: ref.col + (ref.colLock ? 0 : b.col - a.col),
  }));
}

/** One inserted row in the formula's own sheet. S07 handles ranges and sheet identities in full. */
export function insertFormulaRow(source: string, at: number): string {
  return transform(source, ref => ({ row: ref.row + (!ref.prefix && ref.row >= at ? 1 : 0), col: ref.col }));
}

function decimalShift(value: number, places: number): number {
  const [digits, exponent] = value.toString().split('e');
  return Number(`${digits}e${Number(exponent ?? 0) + places}`);
}

function numeric(value: FormulaValue): number {
  if (typeof value === 'number') return value;
  if (typeof value === 'boolean') return Number(value);
  if (value === null) return 0;
  if (typeof value === 'string' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/iu.test(value.trim())) return Number(value);
  return NaN;
}

export function evaluateProbe(source: string, lookup: (address: string) => FormulaValue = () => null): FormulaValue {
  let root: Expression;
  try { root = parse(source).expression; } catch { return error('#VALUE!'); }
  const finite = (number: number) => Number.isFinite(number) ? number || 0 : error('#NUM!');
  function evaluate(node: Expression): FormulaValue {
    switch (node.kind) {
      case 'literal': return node.value;
      case 'reference': {
        if (node.last) return error('#VALUE!');
        const ref = coordinates(node.first.text);
        return lookup(`${ref.prefix}${letters(ref.col)}${ref.row}`);
      }
      case 'sign': {
        const value = evaluate(node.value);
        if (isError(value)) return value;
        const n = numeric(value);
        return Number.isNaN(n) ? error('#VALUE!') : finite(node.operator === '-' ? -n : n);
      }
      case 'binary': {
        const a = evaluate(node.left);
        if (isError(a)) return a;
        const b = evaluate(node.right);
        if (isError(b)) return b;
        const left = numeric(a);
        const right = numeric(b);
        if (Number.isNaN(left) || Number.isNaN(right)) return error('#VALUE!');
        if (node.operator === '/' && right === 0) return error('#DIV/0!');
        if (node.operator === '+') return finite(left + right);
        if (node.operator === '-') return finite(left - right);
        if (node.operator === '*') return finite(left * right);
        return finite(left / right);
      }
      case 'call': {
        const [first, second, third] = node.args;
        if (node.name === 'IF' && first && second && third && node.args.length === 3) {
          const condition = evaluate(first);
          if (isError(condition)) return condition;
          const n = numeric(condition);
          return Number.isNaN(n) ? error('#VALUE!') : evaluate(n === 0 ? third : second);
        }
        if (node.name === 'IFERROR' && first && second && node.args.length === 2) {
          const value = evaluate(first);
          return isError(value) ? evaluate(second) : value;
        }
        if (node.name === 'ABS' && first && node.args.length === 1) {
          const value = evaluate(first);
          if (isError(value)) return value;
          const n = numeric(value);
          return Number.isNaN(n) ? error('#VALUE!') : finite(Math.abs(n));
        }
        if (['ROUND', 'ROUNDUP', 'ROUNDDOWN'].includes(node.name) && first && second && node.args.length === 2) {
          const value = evaluate(first);
          if (isError(value)) return value;
          const digits = evaluate(second);
          if (isError(digits)) return digits;
          if (typeof value !== 'number' || typeof digits !== 'number') return error('#VALUE!');
          const places = Math.trunc(digits);
          if (places > 308 || value === 0) return value;
          if (places < -308) return 0;
          const scaled = decimalShift(Math.abs(value), places);
          if (!Number.isFinite(scaled)) return value;
          const round = node.name === 'ROUNDUP' ? Math.ceil : node.name === 'ROUNDDOWN' ? Math.floor : Math.round;
          return finite(Math.sign(value) * decimalShift(round(scaled), -places));
        }
        if (['IF', 'IFERROR', 'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN'].includes(node.name)) return error('#VALUE!');
        return error('#NAME?');
      }
    }
  }
  return evaluate(root);
}
