import type { CellCache, CellContent, CellDocument } from '../workbook';
import { formatSections, type FormatSection, type FormatToken } from './format-code';
import { explicitFormatLocale, formatLocale } from './format-locales';
import { temporalDisplay } from './temporal';

type Culture = ReturnType<typeof formatLocale>;
export interface FormattedCell { readonly text: string; readonly color?: string }

export function generalNumber(value: number): string {
  const rounded = Number(value.toPrecision(15));
  if (rounded === 0 || Math.abs(rounded) >= 1e-4 && Math.abs(rounded) < 1e15) return String(rounded);
  return rounded.toExponential().replace(/e([+-])(\d+)$/u, (_, sign: string, exponent: string) =>
    'E' + sign + exponent.padStart(2, '0'));
}

function symbol(token: FormatToken, text: string) { return token.kind === 'symbol' && token.text === text; }

function literal(token: FormatToken): string {
  switch (token.kind) {
    case 'pad': return ' ';
    case 'directive': return token.text.startsWith('$') ? token.text.slice(1).split('-')[0]! : '';
    case 'digit': return token.text === '?' ? ' ' : '';
    default: return token.text;
  }
}

const binary = new DataView(new ArrayBuffer(8));

function fixed(value: number, places: number): string {
  if (value >= 1e21) return BigInt(value).toString() + (places ? '.' + '0'.repeat(places) : '');
  if (places <= 100) {
    // toFixed rounds the stored double exactly. Correct true halfway cases to even.
    const text = value.toFixed(places);
    const lastDigit = Number(text.at(-1));
    return lastDigit % 2 && (value * 2 ** places) % 1 === .5 ? text.slice(0, -1) + (lastDigit - 1) : text;
  }
  // A custom code can request more decimal places than toFixed supports.
  binary.setFloat64(0, value);
  const bits = binary.getBigUint64(0);
  const exponent = Number(bits >> 52n);
  const fraction = bits & ((1n << 52n) - 1n);
  const significand = exponent ? fraction + (1n << 52n) : fraction;
  const power = Math.max(1, exponent) - 1023 - 52;
  const numerator = significand * 10n ** BigInt(places);
  const denominator = 1n << BigInt(-power);
  const quotient = numerator / denominator;
  const remainder = numerator % denominator;
  const roundUp = remainder * 2n > denominator || remainder * 2n === denominator && quotient % 2n !== 0n;
  const digits = (quotient + (roundUp ? 1n : 0n)).toString().padStart(places + 1, '0');
  return digits.slice(0, -places) + '.' + digits.slice(-places);
}

function integerMask(tokens: readonly FormatToken[], digits: string, group = ''): string {
  let pending = digits;
  const first = tokens.findIndex(token => token.kind === 'digit');
  const pieces = Array<string>(tokens.length);
  for (let index = tokens.length - 1; index >= 0; index--) {
    const token = tokens[index]!;
    if (token.kind === 'digit') {
      const char = pending.slice(-1);
      pending = pending.slice(0, -1);
      pieces[index] = (index === first ? pending : '') + (char || (token.text === '0' ? '0' : literal(token)));
    } else pieces[index] = symbol(token, ',') ? '' : literal(token);
  }
  if (group) {
    let count = 0;
    for (let index = pieces.length - 1; index >= first && first >= 0; index--) {
      if (tokens[index]?.kind !== 'digit') continue;
      let grouped = '';
      for (const ch of [...pieces[index]!].reverse()) {
        if (/^\d$/u.test(ch) && count++ > 0 && (count - 1) % 3 === 0) grouped = ch + group + grouped;
        else grouped = ch + grouped;
      }
      pieces[index] = grouped;
    }
  }
  return pieces.join('');
}

function splitMantissa(tokens: readonly FormatToken[]) {
  const dot = tokens.findIndex(token => symbol(token, '.'));
  const left = dot < 0 ? tokens : tokens.slice(0, dot);
  const right = dot < 0 ? [] : tokens.slice(dot + 1);
  const lastDigit = (parts: readonly FormatToken[]) => parts.map(token => token.kind).lastIndexOf('digit');
  const lastInteger = lastDigit(left);
  const lastFraction = lastDigit(right);
  const scale = (parts: readonly FormatToken[], index: number) => {
    if (index < 0) return 0;
    let end = index + 1;
    while (parts[end] && symbol(parts[end]!, ',')) end++;
    return end - index - 1;
  };
  const integerScale = scale(left, lastInteger);
  const fractionScale = scale(right, lastFraction);
  return {
    left, right: right.filter((_, index) => index <= lastFraction || index > lastFraction + fractionScale),
    dot: dot >= 0, scale: 1000 ** (integerScale + fractionScale),
    grouping: left.slice(0, lastInteger).some(token => symbol(token, ',')),
  };
}

function decimalNumber(value: number, tokens: readonly FormatToken[], culture: Culture): string {
  const { left, right, dot, scale, grouping } = splitMantissa(tokens);
  const places = right.filter(token => token.kind === 'digit').length;
  const [whole, fraction = ''] = fixed(value / scale, places).split('.') as [string, string?];
  let index = 0;
  const significant = fraction.replace(/0+$/u, '').length;
  const decimals = right.map(token => {
    if (token.kind !== 'digit') return symbol(token, ',') ? '' : literal(token);
    const at = index++;
    return token.text === '0' || at < significant ? fraction[at] ?? '0' : literal(token);
  }).join('');
  const integral = integerMask(left, whole === '0' ? '' : whole, grouping ? culture.group : '');
  const implicitWhole = places > 0 && !left.some(token => token.kind === 'digit') ? whole : '';
  return integral + implicitWhole + (dot ? culture.decimal : '') + decimals;
}

function nearestFraction(value: number, maxDenominator: number): [number, number] {
  let [a, b, c, d] = [0, 1, 1, 0];
  let remaining = value;
  for (;;) {
    const whole = Math.floor(remaining);
    if (b + whole * d > maxDenominator) break;
    [a, b, c, d] = [c, d, a + whole * c, b + whole * d];
    if (remaining === whole) return [c, d];
    remaining = 1 / (remaining - whole);
  }
  const steps = Math.floor((maxDenominator - b) / d);
  const numerator = a + steps * c;
  const denominator = b + steps * d;
  return Math.abs(value - c / d) <= Math.abs(value - numerator / denominator) ? [c, d] : [numerator, denominator];
}

function fractionNumber(value: number, tokens: readonly FormatToken[], slash: number): string | undefined {
  let start = slash;
  while (tokens[start - 1]?.kind === 'digit') start--;
  if (start === slash) return;
  let end = slash + 1;
  const fixedToken = tokens[end];
  const denominator = fixedToken?.kind === 'symbol' && /^\d+$/u.test(fixedToken.text) ? Number(fixedToken.text) : 0;
  if (denominator) end++;
  else while (tokens[end]?.kind === 'digit') end++;
  if (end === slash + 1) return;
  const whole = tokens.slice(0, start);
  const mixed = whole.some(token => token.kind === 'digit');
  const integral = mixed ? Math.floor(value) : 0;
  const width = end - slash - 1;
  const [n, d] = denominator ? [Number(fixed((value - integral) * denominator, 0)), denominator] :
    nearestFraction(value - integral, 10 ** Math.min(width, 7) - 1);
  const numeratorTokens = tokens.slice(start, slash);
  const divisorTokens = tokens.slice(slash + 1, end);
  const numerator = integerMask(numeratorTokens, String(n));
  const divisor = denominator ? String(d) : divisorTokens.every(token => token.text === '?') ?
    String(d).padEnd(width, ' ') : integerMask(divisorTokens, String(d));
  const fraction = numerator + '/' + divisor;
  const showFraction = !mixed || n !== 0 || numeratorTokens.some(token => token.text === '0');
  const leading = integral || !showFraction || n === 0 ? String(integral) : '';
  return integerMask(whole, leading) + (showFraction ? fraction : ' '.repeat(fraction.length)) + tokens.slice(end).map(literal).join('');
}

function numberSection(value: number, section: FormatSection, culture: Culture): string {
  const tokens = section.tokens;
  if (section.date || section.time) return temporalDisplay(value, tokens, culture.tag, literal) ?? generalNumber(value);
  if (tokens.some(token => token.kind === 'general' || symbol(token, '@'))) {
    return tokens.map(token => {
      if (token.kind === 'general' || symbol(token, '@')) return generalNumber(value);
      return token.kind === 'digit' ? token.raw : literal(token);
    }).join('');
  }
  if (tokens.every(token => token.kind === 'directive' && !token.text.startsWith('$'))) return generalNumber(value);
  const sign = value < 0 ? '-' : '';
  const magnitude = Math.abs(value) * (tokens.some(token => symbol(token, '%')) ? 100 : 1);
  if (!Number.isFinite(magnitude)) return generalNumber(value);
  const slash = tokens.findIndex(token => symbol(token, '/'));
  if (slash >= 0) {
    const fraction = fractionNumber(magnitude, tokens, slash);
    if (fraction !== undefined) return sign + fraction;
  }
  const exp = tokens.findIndex(token => token.kind === 'exponent');
  if (exp < 0) return sign + decimalNumber(magnitude, tokens, culture);
  const mantissa = tokens.slice(0, exp);
  const width = Math.max(1, splitMantissa(mantissa).left.filter(token => token.kind === 'digit').length);
  const exponent = magnitude === 0 ? 0 : Math.floor(Number(magnitude.toExponential().split('e')[1]) / width) * width;
  let significand = magnitude === 0 ? 0 : magnitude / 10 ** Math.max(-323, exponent);
  if (exponent < -323) significand *= 10 ** (-323 - exponent);

  const expSign = exponent < 0 ? '-' : tokens[exp]!.text.endsWith('+') ? '+' : '';
  const exponentTokens = tokens.slice(exp + 1);
  const exponentDigits = String(Math.abs(exponent));
  const implicitExponent = exponentTokens.some(token => token.kind === 'digit') ? '' : exponentDigits;
  return sign + decimalNumber(significand, mantissa, culture) + tokens[exp]!.text[0] + expSign +
    implicitExponent + integerMask(exponentTokens, exponentDigits);
}

export function formatNumber(value: number, code?: string): FormattedCell {
  const parsed = code ? formatSections(code) : null;
  if (!parsed || /^general$/iu.test(code ?? '')) return { text: generalNumber(value) };
  // A trailing text section does not change the positive/negative/zero rules.
  const empty = parsed.findIndex(section => section.tokens.length === 0);
  const sections = empty >= 0 ? parsed.slice(0, empty) :
    parsed.length > 1 && parsed.at(-1)?.tokens.some(token => symbol(token, '@')) ? parsed.slice(0, -1) : parsed;
  const conditional = sections.some(section => section.condition);
  const index = conditional ? sections.findIndex(section => !section.condition || section.condition(value)) :
    value < 0 && sections.length > 1 ? 1 : value === 0 && sections.length > 2 ? 2 : 0;
  const section = sections[index];
  if (!section) return { text: generalNumber(value) };
  const culture = explicitFormatLocale(code) ?? formatLocale();
  const unsigned = index === 1 || index === 0 && !!section.condition;
  const formatted = numberSection(unsigned ? Math.abs(value) : value, section, culture);
  return { text: formatted, ...(!conditional && section.color && { color: section.color }) };
}

export function displayCellValue(cell: CellDocument | undefined, source = false): CellContent | CellCache | undefined {
  if (!cell) return undefined;
  if (source) return cell.content;
  if (cell.content.kind === 'python') {
    const block = cell.content.block;
    if (block?.hasResult && block.lastFigure) return { kind: 'text', value: '[plot]' };
    if (block?.hasResult && block.lastValue) return block.lastValue;
    return cell.cache ?? { kind: 'blank' };
  }
  if (cell.content.kind === 'blank' && cell.spillKind === 'formula' && cell.cache) return cell.cache;
  return cell.content.kind === 'formula' ? cell.cache ?? cell.content : cell.content;
}

const formattedCache = new WeakMap<CellDocument, FormattedCell>();
const emptyFormatted: FormattedCell = Object.freeze({ text: '' });

export function formatCellValue(cell: CellDocument | undefined, code?: string): FormattedCell {
  if (!cell) return emptyFormatted;
  const effectiveCode = code !== undefined ? code : cell.style?.numberFormat;
  const useDefault = effectiveCode === cell.style?.numberFormat;
  if (useDefault && Object.isFrozen(cell)) {
    const cached = formattedCache.get(cell);
    if (cached) return cached;
  }
  const value = displayCellValue(cell);
  let formatted: FormattedCell;
  if (!value || value.kind === 'blank' || value.kind === 'formula' || value.kind === 'python') formatted = emptyFormatted;
  else if (value.kind === 'number') formatted = formatNumber(value.value, effectiveCode);
  else formatted = { text: value.kind === 'boolean' ? value.value ? 'TRUE' : 'FALSE' : value.value };
  if (useDefault && Object.isFrozen(cell)) formattedCache.set(cell, formatted);
  return formatted;
}
