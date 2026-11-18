import { parseFormula } from '../../core/formula-references';

export interface TextSelection {
  readonly start: number;
  readonly end: number;
}

/** Find an operand slot in an incomplete draft without interpreting quoted text as syntax. */
export function referenceSlot(source: string, { start, end }: TextSelection): TextSelection | undefined {
  if (!source.startsWith('=') || start < 1 || end < start || end > source.length) return;
  let quote = '', brackets = 0, braces = 0;
  for (let i = 1; i < start; i++) {
    const token = source[i];
    if (quote) {
      if (token === quote) {
        if (i + 1 < start && source[i + 1] === quote) i++;
        else quote = '';
      }
    } else if (token === "'" || token === '"') quote = token;
    else if (token === '[') brackets++;
    else if (token === ']') brackets--;
    else if (token === '{') braces++;
    else if (token === '}') braces--;
  }
  if (quote || brackets || braces || !/[=(,+*/^&<>:;\-]$/u.test(source.slice(0, start).trimEnd())) return;
  if (start === end) {
    const next = source.slice(end).trimStart();
    return !next || /^[),;+*/^&=<>%\-]/u.test(next) ? { start, end } : undefined;
  }
  try {
    const selected = parseFormula('=' + source.slice(start, end)).expression;
    if (selected.kind === 'reference' && selected.reference.start === 1 && selected.reference.end === end - start + 1) {
      return { start, end };
    }
  } catch {
    // A selected part of a draft can be incomplete or be something other than a reference.
  }
}
