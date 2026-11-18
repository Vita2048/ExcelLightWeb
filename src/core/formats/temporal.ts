import { excelDateParts } from '../date-serial-probe';
import { temporalParts, type FormatToken } from './format-code';
import { explicitFormatLocale, formatLocale } from './format-locales';

const dayMs = 86_400_000;
const tickSecond = 10_000_000;
const tickDay = tickSecond * 86_400;
const maxSerial = 2_958_466;
const pad = (value: number, width = 2) => String(value).padStart(width, '0');
const names = new Map<string, Intl.DateTimeFormat>();

function nameOf(date: Date, locale: string, part: 'month' | 'weekday' | 'era', width: 'long' | 'short'): string {
  const key = [locale, part, width].join(':');
  let formatter = names.get(key);
  if (!formatter) {
    formatter = new Intl.DateTimeFormat(locale, { [part]: width, calendar: 'gregory', timeZone: 'UTC' });
    names.set(key, formatter);
  }
  return formatter.formatToParts(date).find(piece => piece.type === part)?.value ?? formatter.format(date);
}

export function temporalDisplay(serial: number, tokens: readonly FormatToken[], locale: string,
  literal: (token: FormatToken) => string): string | undefined {
  // Ordinary Excel dates follow the COM convention for negative serial fractions.
  const elapsed = tokens.some(token => token.kind === 'elapsed');
  const magnitude = Math.abs(serial);
  const whole = Math.trunc(elapsed ? magnitude : serial);
  const fraction = magnitude - Math.floor(magnitude);
  const milliseconds = Math.min(dayMs - 1, Math.round(fraction * dayMs));
  const hours = Math.floor(milliseconds / 3_600_000);
  const minutes = Math.floor(milliseconds / 60_000) % 60;
  const seconds = Math.floor(milliseconds / 1000) % 60;
  if (!elapsed && (whole >= maxSerial || whole < -693593)) return;
  const [year, month, day] = elapsed ? [1900, 1, 1] : whole === 0 ? [1900, 1, 0] : excelDateParts(whole < 0 ? whole + 1 : whole);
  const date = new Date(0);
  date.setUTCFullYear(year, month - 1, whole === 60 ? 28 : day || 1);
  const twelveHour = tokens.some(token => token.kind === 'period');
  const totalSeconds = whole * 86_400 + Math.floor(milliseconds / 1000);
  if (!Number.isFinite(totalSeconds)) return;
  let lastDate = '';
  let prefix = elapsed && serial < 0 ? '-' : '';
  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index]!;
    const { kind, text } = token;
    if (kind === 'period') {
      prefix += text.toLowerCase() === 'a/p' ? text[hours < 12 ? 0 : 2] : hours < 12 ? 'AM' : 'PM';
      continue;
    }
    if (kind === 'elapsed') {
      lastDate = text[0]!;
      const unit = lastDate === 'h' ? 3600 : lastDate === 'm' ? 60 : 1;
      prefix += pad(Math.floor(totalSeconds / unit), text.length);
      continue;
    }
    if (kind === 'symbol' && text === '.' && tokens[index + 1]?.kind === 'digit' && tokens[index + 1]?.text === '0') {
      let places = 0;
      while (tokens[index + 1]?.kind === 'digit' && tokens[index + 1]?.text === '0') { index++; places++; }
      prefix += '.' + pad(milliseconds % 1000, places);
      continue;
    }
    if (kind !== 'date') { prefix += kind === 'digit' ? token.raw : literal(token); continue; }
    const width = text.length;
    const unit = text[0]!;
    switch (unit) {
      case 'g': prefix += nameOf(date, locale, 'era', 'short'); break;
      case 'y': prefix += width < 3 ? pad(year % 100) : pad(year, 4); break;
      case 'd': prefix += width > 2 ? nameOf(date, locale, 'weekday', width === 3 ? 'short' : 'long') : pad(day, width); break;
      case 'h': prefix += pad(twelveHour ? hours % 12 || 12 : hours, width); break;
      case 's': prefix += pad(seconds, width); break;
      case 'm': {
        const next = tokens.slice(index + 1).find(part => part.kind === 'date' || part.kind === 'elapsed');
        if (width <= 2) prefix += pad(lastDate === 'h' || next?.text.startsWith('s') ? minutes : month, width);
        else if (width === 3) prefix += nameOf(date, locale, 'month', 'short');
        else {
          const monthName = nameOf(date, locale, 'month', 'long');
          prefix += width === 5 ? [...monthName][0] : monthName;
        }
        break;
      }
    }
    lastDate = unit;
  }
  return prefix;
}

function entryDate(serial: number, pattern: string) {
  const [year, month, day] = excelDateParts(serial);
  return pattern.replace(/"([^"]*)"|yyyy|MM|dd|M|d/gu, (part, literal: string | undefined) => {
    if (literal !== undefined) return literal;
    if (part === 'yyyy') return pad(year, 4);
    return pad(part[0] === 'M' ? month : day, serial === 60 ? 2 : part.length);
  });
}

export function temporalEntry(serial: number, code?: string, locale?: string): string | undefined {
  const parts = temporalParts(code);
  if ((!parts.date && !parts.time) || serial < 0 || serial >= maxSerial || !Number.isFinite(serial)) return;
  const days = Math.floor(serial);
  let ticks = Math.round((serial - days) * tickDay);
  const ulp = serial === 0 ? 0 : 2 ** (Math.floor(Math.log2(serial)) - 52);
  const tolerance = Math.max(.5, ulp * tickDay * 2);
  for (let quantum = tickSecond; quantum >= 1; quantum /= 10) {
    const rounded = Math.round(ticks / quantum) * quantum;
    if (Math.abs(ticks - rounded) <= tolerance) { ticks = rounded; break; }
  }
  ticks = Math.min(ticks, tickDay - 1);
  const seconds = Math.floor(ticks / tickSecond);
  const clock = Math.floor(seconds / 3600) + ':' + pad(Math.floor(seconds / 60) % 60) + ':' + pad(seconds % 60) +
    (ticks % tickSecond ? '.' + pad(ticks % tickSecond, 7).replace(/0+$/u, '') : '');
  if (days === 0) return parts.time ? clock : undefined;
  const culture = explicitFormatLocale(code) ?? formatLocale(locale);
  return entryDate(days, culture.shortDate) + (parts.time || ticks ? ' ' + clock : '');
}

function parseClock(text: string): number | undefined {
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2})([.,]\d{1,7})?)?(?:\s*([ap])m)?$/iu.exec(text);
  if (!match) return;
  const [, hour, minute, second, fraction, period] = match;
  const h = Number(hour), m = Number(minute), s = Number(second ?? 0);
  if (m > 59 || s > 59 || (period ? h < 1 || h > 12 : h > 23)) return;
  const hours = period ? h % 12 + (period.toLowerCase() === 'p' ? 12 : 0) : h;
  return (hours * 3600 + m * 60 + s + Number(fraction?.replace(',', '.') ?? 0)) / 86_400;
}

function parseDate(text: string, pattern: string, tag: string): [number, number, number] | undefined {
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/u.exec(text);
  if (iso) return [Number(iso[1]), Number(iso[2]), Number(iso[3])];
  const order = [...pattern.replace(/"[^"]*"/gu, '').matchAll(/y+|M+|d+/gu)].map(match => match[0]![0]);
  const split = (input: string) => input.split(/[\s.,/-]+/u).filter(Boolean);
  const numericText = [...pattern.matchAll(/"([^"]*)"/gu)].reduce((value, match) => value.replace(match[1]!, ''), text);
  const fields = split(numericText);
  if (fields.some(field => !/^\d+$/u.test(field))) {
    const normalized = numericText.toLocaleLowerCase(tag);
    for (let month = 0; month < 12; month++) {
      const sample = new Date(Date.UTC(2000, month, 15));
      for (const width of ['long', 'short'] as const) {
        const label = nameOf(sample, tag, 'month', width).toLocaleLowerCase(tag);
        const contextual = new Intl.DateTimeFormat(tag, { month: width, day: 'numeric', timeZone: 'UTC' })
          .formatToParts(sample).find(part => part.type === 'month')?.value.toLocaleLowerCase(tag);
        for (const candidate of new Set([label, contextual])) {
          if (!candidate || !normalized.includes(candidate)) continue;
          const values = split(normalized.replace(candidate, '@'));
          if (values.length !== 3 || !values.includes('@')) continue;
          const dates = values.filter(value => value !== '@');
          if (dates.length !== 2 || dates.some(value => !/^\d+$/u.test(value))) continue;
          const longYear = dates.findIndex(value => value.length === 4);
          const yearIndex = longYear >= 0 ? longYear : order[0] === 'y' ? 0 : 1;
          return [entryYear(dates[yearIndex]!), month + 1, Number(dates[1 - yearIndex])];
        }
      }
    }
  }
  if (fields.length !== 3 || fields.some(field => !/^\d+$/u.test(field))) return;
  const yearText = fields[order.indexOf('y')];
  if (!yearText || yearText.length > 4) return;
  return [entryYear(yearText), Number(fields[order.indexOf('M')]), Number(fields[order.indexOf('d')])];
}

function entryYear(text: string): number {
  const year = Number(text);
  return text.length <= 2 ? year + (year <= 49 ? 2000 : 1900) : year;
}

export function parseTemporalEntry(input: string, code?: string, locale?: string): number | undefined {
  const parts = temporalParts(code);
  if (!parts.date && !parts.time) return;
  const clock = parseClock(input);
  if (clock !== undefined) return clock;
  const match = /^(.+?)(?:[ T](\d{1,2}:\d{2}.*))?$/u.exec(input);
  if (!match) return;
  const culture = explicitFormatLocale(code) ?? formatLocale(locale);
  const time = match[2] === undefined ? 0 : parseClock(match[2]);
  if (time === undefined) return;
  const dateText = match[1]!.replace(/\p{Cf}/gu, '');
  if (/^(?:29[./-]0?2[./-]1900|0?2\/29\/1900)$/u.test(dateText)) return 60 + time;
  const date = parseDate(dateText, culture.shortDate, culture.tag);
  if (!date) return;
  const [year, month, day] = date;
  if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return;
  if (year === 1900 && month === 2 && day === 29) return 60 + time;
  const stamp = Date.UTC(year, month - 1, day);
  if (new Date(stamp).getUTCMonth() !== month - 1) return;
  return (stamp - Date.UTC(1899, 11, 31)) / dayMs + (stamp >= Date.UTC(1900, 2, 1) ? 1 : 0) + time;
}
