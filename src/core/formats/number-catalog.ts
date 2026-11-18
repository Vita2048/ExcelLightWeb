import { formatNumber } from './cell-format';
import { formatSections, temporalParts, withoutLocale } from './format-code';
import { explicitFormatLocale, formatLocale, withFormatLocale } from './format-locales';

export const commaStyle = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
export const accountingStyle = '_(* $#,##0.00_);_(* ($#,##0.00);_(* "-"_);_(@_)';
export const presets = {
  General: 'General', Number: '0.00', Currency: '$#,##0.00', Accounting: accountingStyle, 'Short Date': 'yyyy-mm-dd',
  'Long Date': 'dddd, mmmm d, yyyy', Time: 'hh:mm:ss', Percentage: '0.00%', Fraction: '# ?/?', Scientific: '0.00E+00', Text: '@',
} as const;
export const categories = ['General', 'Number', 'Currency', 'Accounting', 'Date', 'Time', 'Percentage',
  'Fraction', 'Scientific', 'Text', 'Special', 'Custom'] as const;
export type NumberCategory = typeof categories[number];
export type FormatName = keyof typeof presets | 'Special' | 'Custom';

export const currencyChoices = {
  CZK: 'Kč Czech', USD: '$ English (United States)', EUR: '€ Euro (123 €)', GBP: '£ English (United Kingdom)',
  'EUR-SK': '€ Slovak', PLN: 'zł Polish',
} as const;
type Currency = keyof typeof currencyChoices;
const moneySymbol: Record<Currency, string> = { CZK: 'Kč', USD: '$', EUR: '€', GBP: '£', 'EUR-SK': '€', PLN: 'zł' };
export type CurrencyChoice = Currency | 'None';
export const currencies = Object.keys(currencyChoices) as Currency[];

const places = (decimals: number) => decimals > 0 ? '.' + '0'.repeat(Math.min(10, decimals)) : '';

export function currencyFormat(currency: Currency, decimals = 2): string {
  const number = '#,##0' + places(decimals);
  if (currency === 'USD') return '$' + number;
  if (currency === 'GBP') return '"£"' + number;
  return number + '" ' + moneySymbol[currency] + '"';
}

export function numericFormat(category: NumberCategory, decimals: number, grouping: boolean, currency: CurrencyChoice,
  negative: number, locale: string): string {
  const digits = (grouping ? '#,##0' : '0') + places(decimals);
  const money = currency === 'None' ? '#,##0' + places(decimals) : currencyFormat(currency, decimals);
  const positive = category === 'Currency' ? money : digits;
  let code = positive;
  if (category === 'Accounting') code = '_(* ' + money + '_);_(* (' + money + ');_(* "-"_);_(@_)';
  else if (category === 'Percentage') code += '%';
  else if (category === 'Scientific') code += 'E+00';
  else if (negative > 0) {
    const alternate = negative < 2 ? positive : /^(?:cs|sk)-/iu.test(locale) ? '-' + positive : '(' + positive + ')';
    code += ';' + (negative % 2 ? '[Red]' : '') + alternate;
  }
  return withFormatLocale(code, locale);
}

export function numericOptions(code: string, locale: string) {
  const raw = withoutLocale(code);
  const sections = formatSections(raw);
  const tokens = sections?.[0]?.tokens ?? [];
  const body = tokens.filter(token => token.kind === 'digit' || token.kind === 'symbol').map(token => token.text).join('');
  const decimal = /\.([0#?]+)/u.exec(body)?.[1]?.length ?? (raw === 'General' ? 2 : 0);
  const red = sections?.[1]?.color === 'FF0000';
  const explicitNegative = sections?.[1]?.tokens.some(token => token.kind === 'symbol' && /[(-]/u.test(token.text));
  const region: Record<string, Currency> = { 'cs-CZ': 'CZK', 'en-GB': 'GBP', 'sk-SK': 'EUR-SK',
    'pl-PL': 'PLN', 'de-DE': 'EUR', 'fr-FR': 'EUR' };
  const money = raw.includes('€') && locale === 'sk-SK' ? 'EUR-SK' : currencies.find(option => raw.includes(moneySymbol[option]));
  return { decimals: Math.min(decimal, 10), grouping: body.includes('#,##'), currency: money ??
    (formatName(code) === 'Accounting' ? 'None' : region[locale] ?? 'USD') as CurrencyChoice,
  negative: (red ? 1 : 0) + (explicitNegative ? 2 : 0) };
}

export const fractions = [
  ['Up to one digit (1/4)', '# ?/?'], ['Up to two digits (21/25)', '# ??/??'], ['Up to three digits (312/943)', '# ???/???'],
  ['As halves (1/2)', '# ?/2'], ['As quarters (2/4)', '# ?/4'], ['As eighths (4/8)', '# ?/8'],
  ['As sixteenths (8/16)', '# ??/16'], ['As tenths (3/10)', '# ?/10'], ['As hundredths (30/100)', '# ??/100'],
] as const;

const usSpecial = [
  ['Zip Code', '00000'], ['Zip Code + 4', '00000-0000'],
  ['Phone Number', '[<=9999999]###-####;(###) ###-####'], ['Social Security Number', '000-00-0000'],
] as const;
const specials = {
  'cs-CZ': [['PSČ', '000\\ 00'], ['PSČ (bez mezery)', '00000'], ['telefonní číslo', '000\\ 000\\ 000'],
    ['telefonní číslo (dlouhé)', '"+420 "000\\ 000\\ 000']],
  'sk-SK': [['PSČ', '000\\ 00'], ['PSČ (bez medzery)', '00000'], ['Telefónne číslo', '000\\ 000\\ 000'],
    ['Medzinárodné telefónne číslo', '"+421 "000\\ 000\\ 000']],
  'en-US': usSpecial,
  'en-GB': [['Telephone number', '00000\\ 000000'], ['International telephone number', '"+44 "0000\\ 000000']],
  'de-DE': [['Postleitzahl', '00000'], ['Telefonnummer', '0000\\ 0000000']],
  'fr-FR': [['Code postal', '00000'], ['Numéro de téléphone', '00\\ 00\\ 00\\ 00\\ 00']],
  'pl-PL': [['Kod pocztowy', '00\\-000'], ['Numer telefonu', '000\\ 000\\ 000']],
} satisfies Record<string, ReadonlyArray<readonly [string, string]>>;
type SpecialLocale = keyof typeof specials;
export const specialLocales = Object.keys(specials) as SpecialLocale[];
const specialCodes = new Set(Object.values(specials).flatMap(choices => choices.map(([, code]) => code)));

export interface FormatChoice { readonly label: string; readonly code: string }
export function localeForCategory(category: NumberCategory, code: string, fallback: string): string {
  const locale = (explicitFormatLocale(code) ?? formatLocale(fallback)).tag;
  return category === 'Special' && !Object.hasOwn(specials, locale) ? 'en-US' : locale;
}

export function typeChoices(category: NumberCategory, tag: string): readonly FormatChoice[] {
  if (category === 'Fraction') return fractions.map(([label, code]) => ({ label, code }));
  if (category === 'Special') {
    return (specials[tag as SpecialLocale] ?? usSpecial).map(([label, code]) => ({ label, code: withFormatLocale(code, tag) }));
  }
  const time = category === 'Time';
  const locale = formatLocale(tag);
  const codes = time ? [locale.longTime, 'hh:mm', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'mm:ss.0', '[h]:mm:ss'] :
    [locale.shortDate, locale.longDate, 'yyyy-mm-dd', 'd.m', 'd.m.yy', 'dd.mm.yy', 'd-mmm', 'mmm-yy'];
  return codes.map((pattern, index) => {
    const code = withFormatLocale(pattern, tag);
    const sample = formatNumber(time ? .563136574074074 : 40982, code);
    return { label: (index < (time ? 1 : 2) ? '*' : '') + sample.text, code };
  });
}

export function formatName(code?: string): FormatName {
  if (!code || /^general$/iu.test(code)) return 'General';
  if (code === '@') return 'Text';
  const raw = withoutLocale(code);
  if (specialCodes.has(raw)) return 'Special';
  const parts = temporalParts(raw);
  const tokens = formatSections(raw)?.flatMap(section => section.tokens) ?? [];
  if (parts.date) return tokens.some(token => token.kind === 'date' && /dddd|mmmm/u.test(token.text)) ? 'Long Date' : 'Short Date';
  if (parts.time) return 'Time';
  if (tokens.some(token => token.kind === 'symbol' && token.text === '%')) return 'Percentage';
  if (tokens.some(token => token.kind === 'exponent')) return 'Scientific';
  if (tokens.some(token => token.kind === 'symbol' && token.text === '/')) return 'Fraction';
  if (tokens.some(token => token.kind === 'pad' || token.kind === 'fill')) return 'Accounting';
  if (/\$|Kč|€|£|zł/u.test(raw)) return 'Currency';
  if (/^(?:0|#,##0)(?:\.[0#]+)?(?:;.*)?$/u.test(raw) && !/["\\]/u.test(raw)) return 'Number';
  return 'Custom';
}

export function changeDecimals(code: string | undefined, direction: 1 | -1): string {
  const kind = formatName(code);
  if (kind === 'General') return '0' + places(direction > 0 ? 1 : 0);
  if (kind === 'Long Date' || kind === 'Short Date' || kind === 'Time' || kind === 'Text' || kind === 'Fraction' || kind === 'Special') return code!;
  const sections = formatSections(code!);
  if (!sections) return code!;
  return sections.map(({ tokens }) => {
    const first = tokens.findIndex(token => token.kind === 'digit' && token.text !== '?');
    if (first < 0) return tokens.map(token => token.raw).join('');
    const raw = tokens.slice(first).map(token => token.raw).join('');
    const match = /^([0#?][0#?,]*)(?:\.([0#?]*))?/u.exec(raw)!;
    const [, integer, decimal] = match;
    const count = Math.max(0, Math.min(10, (decimal?.length ?? 0) + direction));
    const separator = count === 0 && decimal === '' ? '.' : places(count);
    return tokens.slice(0, first).map(token => token.raw).join('') + integer! + separator + raw.slice(match[0].length);
  }).join(';');
}
