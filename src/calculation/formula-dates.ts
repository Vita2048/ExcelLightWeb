import { excelDateParts } from '../core/date-serial-probe';
import { error, finite, type Scalar } from './formula-values';

const dayMs = 86_400_000;
const epoch = Date.UTC(1899, 11, 31);
const march1900 = Date.UTC(1900, 2, 1);
const lastSerial = 2_958_465;

export function todaySerial(date: Date): number {
  const midnight = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return (midnight - epoch) / dayMs + (midnight >= march1900 ? 1 : 0);
}

function parseDateText(text: string, locale: string): number {
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/u.exec(text.trim());
  let year: number, month: number, day: number;
  if (iso) {
    [, year, month, day] = iso.map(Number) as [number, number, number, number];
  } else {
    const options = { calendar: 'gregory', day: 'numeric', year: 'numeric' } as const;
    const sample = new Date(2001, 0, 23);
    let parts = new Intl.DateTimeFormat(locale, { ...options, month: 'numeric' }).formatToParts(sample);
    const separator = parts.find(part => part.type === 'literal')?.value.trim();
    let fields = separator ? text.trim().split(separator).map(field => field.trim()) : [];
    if (fields.length !== 3 || fields.some(field => !/^\d+$/u.test(field))) {
      const normalized = text.trim().toLocaleLowerCase(locale);
      search: for (const width of ['long', 'short'] as const) {
        const formatter = new Intl.DateTimeFormat(locale, { ...options, month: width });
        for (let month = 0; month < 12; month++) {
          const date = new Date(2001, month, 23);
          const sampleParts = formatter.formatToParts(date);
          const label = sampleParts.find(part => part.type === 'month')?.value.toLocaleLowerCase(locale);
          if (!label || !/\p{L}/u.test(label) || !normalized.includes(label)) continue;
          const candidate = normalized.replace(label, String(month + 1)).split(/[\s.,/-]+/u);
          if (candidate.length !== 3 || candidate.some(field => !/^\d+$/u.test(field))) continue;
          parts = sampleParts;
          fields = candidate;
          break search;
        }
      }
    }
    if (fields.length !== 3 || fields.some(field => !/^\d+$/u.test(field))) return NaN;
    const order = parts.filter(part => part.type === 'year' || part.type === 'month' || part.type === 'day');
    const values = new Map(order.map((part, index) => [part.type, Number(fields[index])]));
    year = values.get('year')!;
    if (fields[order.findIndex(part => part.type === 'year')]!.length <= 2) year += year <= 49 ? 2000 : 1900;
    month = values.get('month')!;
    day = values.get('day')!;
  }
  if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return NaN;
  const stamp = Date.UTC(year, month - 1, day);
  if (new Date(stamp).getUTCMonth() !== month - 1) return NaN;
  return (stamp - epoch) / dayMs + (stamp >= march1900 ? 1 : 0);
}

export function dateFunction(name: string, args: readonly Scalar[], today: number,
  number: (value: Scalar) => number, locale: string): Scalar {
  if (name === 'TODAY') return args.length === 0 ? finite(today) : error('#VALUE!');
  const expected = name === 'DATE' ? 3 : 1;
  if (args.length !== expected) return error('#VALUE!');
  const numbers = args.map(value => {
    const numeric = number(value);
    return name !== 'DATE' && typeof value === 'string' && Number.isNaN(numeric) ? parseDateText(value, locale) : numeric;
  });
  if (numbers.some(Number.isNaN)) return error('#VALUE!');
  if (numbers.some(value => !Number.isFinite(value))) return error('#NUM!');
  if (name === 'DATE') {
    const [year, month, day] = numbers.map(Math.trunc) as [number, number, number];
    if (year < 0 || year > 9999) return error('#NUM!');
    const fullYear = year < 1900 ? year + 1900 : year;
    const stamp = Date.UTC(fullYear, month - 1, 1);
    const serial = (stamp - epoch) / dayMs + (stamp >= march1900 ? 1 : 0) + day - 1;
    return Number.isFinite(serial) && serial >= 0 && serial <= lastSerial ? serial : error('#NUM!');
  }
  const serial = numbers[0]!;
  if (serial < 0 || serial >= lastSerial + 1) return error('#NUM!');
  const parts = serial < 1 ? [1900, 1, 0] : excelDateParts(serial);
  return parts[name === 'YEAR' ? 0 : name === 'MONTH' ? 1 : 2]!;
}
