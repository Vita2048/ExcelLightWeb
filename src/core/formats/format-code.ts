import { documentLimits } from '../workbook';

type TokenKind = 'digit' | 'symbol' | 'literal' | 'directive' | 'pad' | 'fill' | 'date' | 'elapsed' | 'period' | 'exponent' | 'general';
export interface FormatToken { readonly kind: TokenKind; readonly text: string; readonly raw: string }
export interface FormatSection {
  readonly tokens: readonly FormatToken[];
  readonly date: boolean;
  readonly time: boolean;
  readonly condition?: (value: number) => boolean;
  readonly color?: string;
}

const colors: Readonly<Record<string, string>> = {
  black: '000000', blue: '0000FF', cyan: '00FFFF', green: '008000', magenta: 'FF00FF', red: 'FF0000', white: 'FFFFFF', yellow: 'FFFF00',
};
const lexical = /"[^"]*"|\\.|[_*].|\[[^\]]*\]|general|am\/pm|a\/p|e[+-]|y+|m+|d+|h+|s+|g+|[0#?]|[1-9]\d*|[-+ /:!^&'~{}<>=(),.%$@;]|[^\x00-\x7f]/iyu;
const cache = new Map<string, readonly FormatSection[] | null>();

function parseFormat(code: string): readonly FormatSection[] | null {
  if (!code || code.length > documentLimits.numberFormatLength || /[\x00-\x1f]/u.test(code)) return null;
  const sections: FormatToken[][] = [[]];
  while (lexical.lastIndex < code.length) {
    const raw = lexical.exec(code)?.[0];
    if (!raw) return null;
    if (raw === ';') {
      if (sections.length === 4) return null;
      sections.push([]);
      continue;
    }
    const add = (kind: TokenKind, text = raw) => sections.at(-1)!.push({ kind, text, raw });
    const first = raw[0]!;
    const lower = raw.toLowerCase();
    if (first === '"' || first === '\\') add('literal', raw.slice(1, first === '"' ? -1 : undefined));
    else if (first === '_') add('pad', raw.slice(1));
    else if (first === '*') add('fill', raw.slice(1));
    else if (first === '[') {
      const body = raw.slice(1, -1);
      if (/^[hms]+$/iu.test(body)) add('elapsed', body.toLowerCase());
      else add('directive', body);
    } else if (/^[0#?]$/u.test(raw)) add('digit');
    else if (lower === 'general') add('general');
    else if (lower === 'am/pm' || lower === 'a/p') add('period');
    else if (/^e[+-]$/u.test(lower)) add('exponent');
    else if (/^[ymdhsg]/u.test(lower)) add('date', lower);
    else add('symbol');
  }
  const parsed = sections.map(tokens => {
    const dateTokens = tokens.filter(token => token.kind === 'date' || token.kind === 'elapsed').map(token => token.text[0]);
    const time = tokens.some(token => token.kind === 'elapsed' || token.kind === 'period') ||
      dateTokens.includes('h') || dateTokens.includes('s');
    const date = dateTokens.includes('y') || dateTokens.includes('d') || !time && dateTokens.includes('m');
    const directives = tokens.filter(token => token.kind === 'directive');
    const color = directives.map(token => colors[token.text.toLowerCase()]).find(Boolean);
    const match = directives.map(token =>
      /^(<>|<=|>=|[<>=])([+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)$/iu.exec(token.text)).find(Boolean);
    let condition: FormatSection['condition'];
    if (match) {
      const [, operator, operand] = match;
      const limit = Number(operand);
      switch (operator) {
        case '<': condition = value => value < limit; break;
        case '<=': condition = value => value <= limit; break;
        case '>': condition = value => value > limit; break;
        case '>=': condition = value => value >= limit; break;
        case '=': condition = value => value === limit; break;
        case '<>': condition = value => value !== limit; break;
      }
    }
    return { tokens, date, time, ...(condition && { condition }), ...(color && { color }) };
  });
  if (parsed.some(({ tokens, date, time }) => !date && !time &&
    tokens.some(token => token.kind === 'digit') && tokens.some(token => token.kind === 'symbol' && token.text === ':'))) return null;
  return parsed;
}

export function formatSections(code: string): readonly FormatSection[] | null {
  if (!cache.has(code)) {
    lexical.lastIndex = 0;
    if (cache.size >= 512) cache.delete(cache.keys().next().value!);
    cache.set(code, parseFormat(code));
  }
  return cache.get(code)!;
}
export function isNumberFormat(code: string): boolean { return formatSections(code) !== null; }

const localeDirective = /^\$-([0-9a-f]+)$/iu;

export function localeId(code: string | undefined): number | undefined {
  for (const section of formatSections(code ?? '') ?? []) {
    for (const token of section.tokens) {
      const id = token.kind === 'directive' && localeDirective.exec(token.text)?.[1];
      if (id) return Number.parseInt(id, 16);
    }
  }
}

export function withoutLocale(code: string): string {
  const sections = formatSections(code);
  return sections ? sections.map(section => section.tokens.filter(token =>
    token.kind !== 'directive' || !localeDirective.test(token.text)).map(token => token.raw).join('')).join(';') : code;
}

export function temporalParts(code: string | undefined) {
  const sections = formatSections(code ?? '') ?? [];
  return { date: sections.some(section => section.date), time: sections.some(section => section.time) };
}
