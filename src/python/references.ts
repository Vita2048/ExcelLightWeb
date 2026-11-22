import type { PythonBlock } from '../core/workbook';

// Port of the desktop PythonReferences lexer. Only the deliberately small xl call
// grammar is recognized. Strings and comments are lexed so unrelated text is never a
// dependency, and editor drafts with incomplete or unsupported syntax stay saveable.
export interface PythonReference {
  readonly start: number;
  readonly length: number;
  readonly reference: string;
  readonly headers: boolean;
}

interface Token {
  readonly text: string;
  readonly start: number;
  readonly length: number;
  readonly value?: string | null;
}

const stringPrefixes = new Set(['r', 'u', 'b', 'f', 'fr', 'rf', 'br', 'rb']);
const identifierStart = /^\p{L}/u;
const identifierPart = /^[\p{L}\p{N}_]/u;
const referencePattern = /^=(?:(?:'(?<quoted>(?:[^']|'')+)'|(?<sheet>[^!']+))!)?(?<first>\$?[A-Za-z]+\$?[0-9]+)(?::(?<last>\$?[A-Za-z]+\$?[0-9]+))?$/u;
const namePattern = /^[A-Za-z_][A-Za-z_0-9]{0,254}$/u;
const addressPattern = /^[A-Za-z]+[0-9]+$/u;
const r1c1Pattern = /^R[0-9]*C[0-9]*$/iu;

function invalid(message = 'Use xl("A1:B3", headers=True/False) with a literal A1, workbook name or TableName[#All] reference and optional literal headers bool. Dynamic references, aliases and shadowing xl are unsupported.'): Error {
  return new Error(message);
}

export function isPythonIdentifier(value: string): boolean {
  return namePattern.test(value) && !addressPattern.test(value) && !r1c1Pattern.test(value) &&
    value.toLowerCase() !== 'r' && value.toLowerCase() !== 'c';
}

export function pythonNameKey(reference: string): string {
  if (reference.toLowerCase().endsWith('[#all]')) return reference.slice(0, -6);
  if (reference.toLowerCase().endsWith('[#data]')) return reference.slice(0, -7);
  return reference;
}

export function isPythonNamedReference(reference: string): boolean {
  return isPythonIdentifier(reference) ||
    reference.toLowerCase().endsWith('[#all]') && isPythonIdentifier(reference.slice(0, -6)) ||
    reference.toLowerCase().endsWith('[#data]') && isPythonIdentifier(reference.slice(0, -7));
}

/** Formula-spelling normalization for a Python binding or literal xl reference. */
export function normalizePythonReference(reference: string): string {
  return referencePattern.test(reference) ? reference : '=#REF!';
}

function hexValue(text: string): number | undefined {
  if (!/^[0-9a-f]+$/iu.test(text)) return undefined;
  const value = Number.parseInt(text, 16);
  return Number.isSafeInteger(value) && value >= 0 && value <= 0x10ffff && !(value >= 0xd800 && value <= 0xdfff) ? value : undefined;
}

function lex(source: string): Token[] {
  const tokens: Token[] = [];
  for (let i = 0; i < source.length;) {
    const character = source[i]!;
    if (/\s/u.test(character)) { i++; continue; }
    if (character === '#') { while (i < source.length && source[i] !== '\n') i++; continue; }
    if (character === '\\' && (source[i + 1] === '\r' || source[i + 1] === '\n')) { i += 2; continue; }
    const start = i;
    let prefix = '';
    if (identifierStart.test(character) || character === '_') {
      while (i < source.length && (identifierPart.test(source[i]!) || source[i] === '_')) i++;
      prefix = source.slice(start, i);
      const quote = source[i];
      if ((quote !== '\'' && quote !== '"') || !stringPrefixes.has(prefix.toLowerCase())) {
        tokens.push({ text: prefix, start, length: i - start });
        continue;
      }
    }
    if (source[i] === '\'' || source[i] === '"') {
      const quote = source[i++]!;
      const triple = source[i] === quote && source[i + 1] === quote;
      if (triple) i += 2;
      const value: string[] = [];
      let closed = false;
      const lower = prefix.toLowerCase();
      let supported = !lower.includes('f') && !lower.includes('b');
      while (i < source.length) {
        if (source[i] === quote && (!triple || source[i + 1] === quote && source[i + 2] === quote)) {
          i += triple ? 3 : 1;
          closed = true;
          break;
        }
        const current = source[i++]!;
        if (current !== '\\' || i >= source.length) { value.push(current); continue; }
        const escape = source[i++]!;
        if (lower.includes('r')) { value.push(current, escape); continue; }
        if (escape === '\\' || escape === '\'' || escape === '"') { value.push(escape); continue; }
        if (escape === 'x' || escape === 'u' || escape === 'U') {
          const digits = escape === 'x' ? 2 : escape === 'u' ? 4 : 8;
          const code = i + digits <= source.length ? hexValue(source.slice(i, i + digits)) : undefined;
          if (code !== undefined) {
            value.push(String.fromCodePoint(code));
            i += digits;
            continue;
          }
          supported = false;
          value.push(escape);
          continue;
        }
        supported = false;
        value.push(escape);
      }
      tokens.push({ text: 'string', start, length: i - start, value: closed && supported ? value.join('') : null });
      continue;
    }
    tokens.push({ text: character, start, length: 1 });
    i++;
  }
  return tokens;
}

export interface PythonReferenceSlot {
  /** Character offsets inside the string literal, excluding its quotes. */
  readonly start: number;
  readonly end: number;
  readonly quote: string;
  readonly raw: boolean;
  readonly triple: boolean;
  readonly literalStart: number;
  readonly literalEnd: number;
}

/**
 * Find a caret inside the first literal argument of a real xl(...) call. The shared
 * lexer keeps comments, outer strings, f-strings and escaped quotes from looking like
 * source code, including while the surrounding editor draft is otherwise incomplete.
 */
export function pythonReferenceSlotAt(source: string, offset: number): PythonReferenceSlot | undefined {
  if (!Number.isInteger(offset) || offset < 0 || offset > source.length) return undefined;
  const tokens = lex(source);
  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index]!;
    if (token.text !== 'string' || token.value == null || index < 2 ||
      tokens[index - 1]!.text !== '(' || tokens[index - 2]!.text !== 'xl' ||
      index > 2 && tokens[index - 3]!.text === '.') continue;
    const literalStart = token.start;
    const literalEnd = token.start + token.length;
    let quoteStart = literalStart;
    while (source[quoteStart] !== '\'' && source[quoteStart] !== '"' && quoteStart < literalEnd) quoteStart++;
    const quote = source[quoteStart];
    if ((quote !== '\'' && quote !== '"') || quoteStart >= literalEnd) continue;
    const triple = source.slice(quoteStart, quoteStart + 3) === quote.repeat(3);
    const quoteWidth = triple ? 3 : 1;
    const start = quoteStart + quoteWidth;
    const end = literalEnd - quoteWidth;
    if (offset < start || offset > end || !triple && /[\r\n]/u.test(source.slice(start, end))) continue;
    const prefix = source.slice(literalStart, quoteStart).toLowerCase();
    return { start, end, quote, raw: prefix.includes('r'), triple, literalStart, literalEnd };
  }
  return undefined;
}

/**
 * Discover literal xl calls in source order. Throws for dynamic references, aliases,
 * unsupported headers and shadowing xl; unrelated strings and comments are ignored.
 */
export function discoverPythonReferences(source: string): PythonReference[] {
  const tokens = lex(source);
  const found: PythonReference[] = [];
  for (let i = 0; i < tokens.length; i++) {
    if (tokens[i]!.text !== 'xl') continue;
    const at = i;
    const next = (value: string) => {
      i++;
      return i < tokens.length && tokens[i]!.text === value;
    };
    if (at > 0 && tokens[at - 1]!.text === '.' || !next('(') || ++i >= tokens.length || tokens[i]!.value == null) {
      throw invalid();
    }
    const literal = tokens[i]!;
    let headers = false;
    if (++i < tokens.length && tokens[i]!.text === ',') {
      if (i + 1 < tokens.length && tokens[i + 1]!.text === ')') i++;
      else {
        if (!next('headers') || !next('=') || ++i >= tokens.length ||
            tokens[i]!.text !== 'True' && tokens[i]!.text !== 'False') throw invalid();
        headers = tokens[i]!.text === 'True';
        i++;
        if (i < tokens.length && tokens[i]!.text === ',') i++;
      }
    }
    if (i >= tokens.length || tokens[i]!.text !== ')') throw invalid();
    const reference = literal.value!;
    if (reference !== '#REF!' && !isPythonNamedReference(reference) &&
        normalizePythonReference('=' + reference) === '=#REF!') throw invalid();
    found.push({ start: literal.start, length: literal.length, reference, headers });
    if (found.length > 100) throw new Error('At most 100 literal xl calls are supported per cell.');
  }
  return found;
}

/**
 * Rewrite only actual literal xl references through a formula-source transform.
 * Incomplete editor drafts are returned unchanged rather than partially rewritten.
 */
export function transformPythonReferences(source: string, transform: (reference: string) => string): string {
  let references: PythonReference[];
  try { references = discoverPythonReferences(source); }
  catch { return source; }
  for (let index = references.length - 1; index >= 0; index--) {
    const reference = references[index]!;
    if (reference.reference === '#REF!' || isPythonNamedReference(reference.reference)) continue;
    const next = normalizePythonReference(transform('=' + reference.reference)).slice(1);
    if (next !== reference.reference) {
      source = source.slice(0, reference.start) + JSON.stringify(next) + source.slice(reference.start + reference.length);
    }
  }
  return source;
}

/** Rename workbook name/table references, preserving a #All/#Data selector. */
export function renamePythonReferences(source: string, oldName: string, newName: string): string {
  let references: PythonReference[];
  try { references = discoverPythonReferences(source); }
  catch { return source; }
  for (let index = references.length - 1; index >= 0; index--) {
    const reference = references[index]!;
    if (pythonNameKey(reference.reference).toLowerCase() !== oldName.toLowerCase()) continue;
    const selector = reference.reference.toLowerCase().endsWith('[#all]') ? '[#All]' :
      reference.reference.toLowerCase().endsWith('[#data]') ? '[#Data]' : '';
    const next = newName === '#REF!' ? '#REF!' : newName + selector;
    source = source.slice(0, reference.start) + JSON.stringify(next) + source.slice(reference.start + reference.length);
  }
  return source;
}

export interface PythonCellContent {
  readonly kind: 'python';
  readonly blockId: string;
  readonly source: string;
  readonly block?: PythonBlock;
}

/** Rewrite both the mirrored source on the cell and its block, keeping their identity in step. */
export function rewritePythonContent<C extends PythonCellContent>(content: C,
  transform: (reference: string) => string): C {
  const block = content.block && transformPythonBlock(content.block, transform);
  const source = transformPythonReferences(content.source, transform);
  if (block === content.block && source === content.source) return content;
  return { ...content, blockId: block?.id ?? content.blockId, source: block?.source ?? source,
    ...(block && { block }) };
}

export function renamePythonContent<C extends PythonCellContent>(content: C, oldName: string, newName: string): C {
  const block = content.block && renamePythonBlock(content.block, oldName, newName);
  const source = renamePythonReferences(content.source, oldName, newName);
  if (block === content.block && source === content.source) return content;
  return { ...content, blockId: block?.id ?? content.blockId, source: block?.source ?? source,
    ...(block && { block }) };
}

export function transformPythonBlock(block: PythonBlock, transform: (reference: string) => string): PythonBlock {
  const source = transformPythonReferences(block.source, transform);
  const bindings = block.bindings.map(binding => binding.reference === '=#REF!' ? binding :
    { ...binding, reference: normalizePythonReference(transform(binding.reference)) });
  const changed = source !== block.source || bindings.some((binding, index) => binding !== block.bindings[index]);
  return changed ? { ...block, source, bindings } : block;
}

export function renamePythonBlock(block: PythonBlock, oldName: string, newName: string): PythonBlock {
  const source = renamePythonReferences(block.source, oldName, newName);
  return source === block.source ? block : { ...block, source };
}
