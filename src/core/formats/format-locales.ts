import { localeId, withoutLocale } from './format-code';
import desktopLocales from './locales.json';

interface FormatLocale {
  readonly id: number;
  readonly tag: string;
  readonly shortDate: string;
  readonly longDate: string;
  readonly longTime: string;
  readonly group: string;
  readonly decimal: string;
}

export const formatLocales: readonly FormatLocale[] = desktopLocales.map(row => {
  const [id, tag, shortDate, longDate, longTime, group, decimal] = row as [number, string, string, string, string, string, string, string];
  return { id, tag, shortDate, longDate, longTime, group, decimal };
});
const byId = new Map(formatLocales.map(locale => [locale.id, locale]));
const byTag = new Map(formatLocales.map(locale => [locale.tag.toLowerCase(), locale]));
const defaultLocale = byId.get(0x409)!;

export function explicitFormatLocale(code?: string) { return byId.get(localeId(code) ?? -1); }
export function formatLocale(tag?: string) { return byTag.get(tag?.toLowerCase() ?? '') ?? defaultLocale; }
export function withFormatLocale(code: string, tag: string) {
  return '[$-' + formatLocale(tag).id.toString(16).toUpperCase() + ']' + withoutLocale(code);
}
