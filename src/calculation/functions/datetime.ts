import { excelDateParts } from '../../core/date-serial-probe';
import type { FormulaExpression } from '../../core/formula-references';
import { dateFunction, dateToSerial, parseDateText, parseTimeText } from '../formula-dates';
import { arrayValues, error, finite, isArray, isError, isOmitted, isRange, type FormulaError, type Scalar } from '../formula-values';
import { evaluateScalarArgs, logicalValue } from './helpers';
import type { FunctionContext, FunctionSignature, FunctionSpec } from './types';

const datePart: FunctionSignature = [['serial_number']];
const lastSerial = 2_958_465;

function dateSpec(
  name: 'DATE' | 'YEAR' | 'MONTH' | 'DAY' | 'TODAY',
  minArgs: number,
  maxArgs: number,
  signature: FunctionSignature,
  volatile = false,
): FunctionSpec {
  return {
    name,
    category: 'datetime',
    minArgs,
    maxArgs,
    ...(volatile && { volatile: true }),
    signature,
    evaluate(args, ctx) {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      return dateFunction(name, values, ctx.today, ctx.number, ctx.locale);
    },
  };
}

function parseSerial(raw: Scalar, ctx: FunctionContext): number | FormulaError {
  if (isError(raw)) return raw;
  const n = ctx.number(raw);
  if (Number.isFinite(n)) return n < 0 || n >= lastSerial + 1 ? error('#NUM!') : n;
  if (typeof raw === 'string') {
    const dt = parseDateText(raw, ctx.locale);
    if (Number.isFinite(dt)) return dt;
    const tm = parseTimeText(raw);
    if (Number.isFinite(tm)) return tm;
    const space = raw.trim().indexOf(' ');
    if (space > 0) {
      const trimmed = raw.trim();
      const d = parseDateText(trimmed.slice(0, space), ctx.locale);
      const t = parseTimeText(trimmed.slice(space + 1));
      if (Number.isFinite(d) && Number.isFinite(t)) return d + t;
    }
  }
  return error('#VALUE!');
}

function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

function isLeapYear(year: number): boolean {
  return year % 400 === 0 || (year % 4 === 0 && year % 100 !== 0);
}

function excelDow(serial: number): number {
  // 0=Sun, 1=Mon, ..., 6=Sat (Excel 1900 date system: serial 1 is Sunday)
  return ((Math.trunc(serial) - 1) % 7 + 7) % 7;
}

function parseWeekendMask(expr: FormulaExpression | undefined, ctx: FunctionContext, rejectAllOff = false): boolean[] | FormulaError {
  if (!expr) return [true, false, false, false, false, false, true]; // 0=Sun..6=Sat (Sat & Sun off)
  const raw = ctx.evaluate(expr);
  if (isOmitted(raw)) return [true, false, false, false, false, false, true];
  const s = ctx.scalar(raw);
  if (isError(s)) return s;
  if (typeof s === 'string' && /^[01]{7}$/u.test(s)) {
    // 7-char string is Mon..Sun -> convert to Sun..Sat
    const mask = [s[6] === '1', s[0] === '1', s[1] === '1', s[2] === '1', s[3] === '1', s[4] === '1', s[5] === '1'];
    return rejectAllOff && mask.every(Boolean) ? error('#VALUE!') : mask;
  }
  if (typeof s === 'string') return error('#VALUE!');
  const code = Math.trunc(ctx.number(s));
  if (!Number.isFinite(code)) return error('#VALUE!');
  const mask = [false, false, false, false, false, false, false];
  switch (code) {
    case 1: mask[6] = true; mask[0] = true; break;
    case 2: mask[0] = true; mask[1] = true; break;
    case 3: mask[1] = true; mask[2] = true; break;
    case 4: mask[2] = true; mask[3] = true; break;
    case 5: mask[3] = true; mask[4] = true; break;
    case 6: mask[4] = true; mask[5] = true; break;
    case 7: mask[5] = true; mask[6] = true; break;
    case 11: mask[0] = true; break;
    case 12: mask[1] = true; break;
    case 13: mask[2] = true; break;
    case 14: mask[3] = true; break;
    case 15: mask[4] = true; break;
    case 16: mask[5] = true; break;
    case 17: mask[6] = true; break;
    default: return error('#NUM!');
  }
  return mask;
}

function parseHolidays(expr: FormulaExpression | undefined, ctx: FunctionContext): Set<number> | FormulaError {
  const holidays = new Set<number>();
  if (!expr) return holidays;
  const val = ctx.evaluate(expr, true);
  if (isOmitted(val)) return holidays;
  if (isError(val)) return val;
  const add = (item: Scalar): FormulaError | undefined => {
    if (isError(item)) return item;
    if (item === null) return undefined;
    const serial = parseSerial(item, ctx);
    if (isError(serial)) return serial;
    holidays.add(Math.trunc(serial));
    return undefined;
  };
  if (isRange(val)) {
    for (const r of val) for (const item of ctx.values(r)) {
      const err = add(item);
      if (err) return err;
    }
  } else if (isArray(val)) {
    for (const item of arrayValues(val)) {
      const err = add(item);
      if (err) return err;
    }
  } else {
    const err = add(ctx.scalar(val));
    if (err) return err;
  }
  return holidays;
}

function isoWeekInfo(serial: number): number {
  const [y, m, d] = excelDateParts(Math.trunc(serial));
  const date = new Date(Date.UTC(y, m - 1, d));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil(((date.getTime() - yearStart.getTime()) / 86_400_000 + 1) / 7);
}

function weekdayNumber(serial: number, returnType: number): number | FormulaError {
  const dow = excelDow(serial); // 0=Sun..6=Sat
  switch (returnType) {
    case 1: return dow + 1;
    case 2: return ((dow + 6) % 7) + 1;
    case 3: return (dow + 6) % 7;
    default:
      if (returnType >= 11 && returnType <= 17) {
        const startDow = (returnType - 10) % 7; // 11->1(Mon), 17->0(Sun)
        return ((dow - startDow + 7) % 7) + 1;
      }
      return error('#NUM!');
  }
}

export const datetimeFunctions: readonly FunctionSpec[] = [
  dateSpec('DATE', 3, 3, [['year', 'month', 'day']]),
  dateSpec('YEAR', 1, 1, datePart),
  dateSpec('MONTH', 1, 1, datePart),
  dateSpec('DAY', 1, 1, datePart),
  dateSpec('TODAY', 0, 0, [[]], true),
  {
    name: 'NOW',
    category: 'datetime',
    minArgs: 0,
    maxArgs: 0,
    volatile: true,
    signature: [[]],
    evaluate: (args, ctx) => args.length === 0 ? finite(ctx.now) : error('#VALUE!'),
  },
  {
    name: 'TIME',
    category: 'datetime',
    minArgs: 3,
    maxArgs: 3,
    signature: [['hour', 'minute', 'second']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 3) return error('#VALUE!');
      const h = Math.trunc(ctx.number(values[0]!));
      const m = Math.trunc(ctx.number(values[1]!));
      const s = Math.trunc(ctx.number(values[2]!));
      if (!Number.isFinite(h) || !Number.isFinite(m) || !Number.isFinite(s)) return error('#VALUE!');
      const totalSeconds = h * 3600 + m * 60 + s;
      if (totalSeconds < 0 || h > 32767 || m > 32767 || s > 32767) return error('#NUM!');
      return (totalSeconds % 86_400) / 86_400;
    },
  },
  {
    name: 'HOUR',
    category: 'datetime',
    minArgs: 1,
    maxArgs: 1,
    signature: datePart,
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const serial = parseSerial(values[0]!, ctx);
      if (isError(serial)) return serial;
      const secs = Math.round((serial - Math.floor(serial)) * 86_400) % 86_400;
      return Math.floor(secs / 3600);
    },
  },
  {
    name: 'MINUTE',
    category: 'datetime',
    minArgs: 1,
    maxArgs: 1,
    signature: datePart,
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const serial = parseSerial(values[0]!, ctx);
      if (isError(serial)) return serial;
      const secs = Math.round((serial - Math.floor(serial)) * 86_400) % 86_400;
      return Math.floor((secs % 3600) / 60);
    },
  },
  {
    name: 'SECOND',
    category: 'datetime',
    minArgs: 1,
    maxArgs: 1,
    signature: datePart,
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const serial = parseSerial(values[0]!, ctx);
      if (isError(serial)) return serial;
      const secs = Math.round((serial - Math.floor(serial)) * 86_400) % 86_400;
      return secs % 60;
    },
  },
  {
    name: 'DATEVALUE',
    category: 'datetime',
    minArgs: 1,
    maxArgs: 1,
    signature: [['date_text']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const raw = values[0];
      if (typeof raw !== 'string') return error('#VALUE!');
      const trimmed = raw.trim();
      const dt = parseDateText(trimmed, ctx.locale);
      if (Number.isFinite(dt)) return Math.trunc(dt);
      const space = trimmed.indexOf(' ');
      if (space > 0) {
        const d = parseDateText(trimmed.slice(0, space), ctx.locale);
        const t = parseTimeText(trimmed.slice(space + 1));
        if (Number.isFinite(d) && Number.isFinite(t)) return Math.trunc(d);
      }
      return error('#VALUE!');
    },
  },
  {
    name: 'TIMEVALUE',
    category: 'datetime',
    minArgs: 1,
    maxArgs: 1,
    signature: [['time_text']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const raw = values[0];
      if (typeof raw !== 'string') return error('#VALUE!');
      const trimmed = raw.trim();
      const tm = parseTimeText(trimmed);
      if (Number.isFinite(tm)) return tm;
      const space = trimmed.indexOf(' ');
      if (space > 0) {
        const d = parseDateText(trimmed.slice(0, space), ctx.locale);
        const t = parseTimeText(trimmed.slice(space + 1));
        if (Number.isFinite(d) && Number.isFinite(t)) return t;
      }
      return error('#VALUE!');
    },
  },
  {
    name: 'EDATE',
    category: 'datetime',
    minArgs: 2,
    maxArgs: 2,
    signature: [['start_date', 'months']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 2) return error('#VALUE!');
      const start = parseSerial(values[0]!, ctx);
      if (isError(start)) return start;
      const months = Math.trunc(ctx.number(values[1]!));
      if (!Number.isFinite(months)) return error('#VALUE!');
      const s = Math.trunc(start);
      const [y, m, d] = s < 1 ? [1900, 1, 0] : excelDateParts(s);
      const totalMonths = (y * 12 + (m - 1)) + months;
      const targetYear = Math.floor(totalMonths / 12);
      const targetMonth = (totalMonths % 12 + 12) % 12 + 1;
      if (targetYear < 1900 || targetYear > 9999) return error('#NUM!');
      const clampedDay = Math.min(Math.max(d, 1), daysInMonth(targetYear, targetMonth));
      return dateToSerial(targetYear, targetMonth, clampedDay);
    },
  },
  {
    name: 'EOMONTH',
    category: 'datetime',
    minArgs: 2,
    maxArgs: 2,
    signature: [['start_date', 'months']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 2) return error('#VALUE!');
      const start = parseSerial(values[0]!, ctx);
      if (isError(start)) return start;
      const months = Math.trunc(ctx.number(values[1]!));
      if (!Number.isFinite(months)) return error('#VALUE!');
      const s = Math.trunc(start);
      const [y, m] = s < 1 ? [1900, 1] : excelDateParts(s);
      const totalMonths = (y * 12 + (m - 1)) + months;
      const targetYear = Math.floor(totalMonths / 12);
      const targetMonth = (totalMonths % 12 + 12) % 12 + 1;
      if (targetYear < 1900 || targetYear > 9999) return error('#NUM!');
      return dateToSerial(targetYear, targetMonth, daysInMonth(targetYear, targetMonth));
    },
  },
  {
    name: 'DATEDIF',
    category: 'datetime',
    minArgs: 3,
    maxArgs: 3,
    signature: [['start_date', 'end_date', 'unit']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 3) return error('#VALUE!');
      const start = parseSerial(values[0]!, ctx);
      if (isError(start)) return start;
      const end = parseSerial(values[1]!, ctx);
      if (isError(end)) return end;
      const s = Math.trunc(start);
      const e = Math.trunc(end);
      if (s > e) return error('#NUM!');
      const unit = ctx.text(values[2]!).trim().toUpperCase();
      if (unit === 'D') return e - s;
      const [y1, m1, d1] = s < 1 ? [1900, 1, 0] : excelDateParts(s);
      const [y2, m2, d2] = e < 1 ? [1900, 1, 0] : excelDateParts(e);
      if (unit === 'Y') {
        let years = y2 - y1;
        if (m2 < m1 || (m2 === m1 && d2 < d1)) years--;
        return years;
      }
      if (unit === 'M') {
        let months = (y2 - y1) * 12 + (m2 - m1);
        if (d2 < d1) months--;
        return months;
      }
      if (unit === 'MD') {
        if (d2 >= d1) return d2 - d1;
        const prevMonth = m2 === 1 ? 12 : m2 - 1;
        const prevYear = m2 === 1 ? y2 - 1 : y2;
        return Math.max(0, daysInMonth(prevYear, prevMonth) - d1 + d2);
      }
      if (unit === 'YM') {
        let months = (m2 - m1 + 12) % 12;
        if (d2 < d1) months = (months + 11) % 12;
        return months;
      }
      if (unit === 'YD') {
        const useSameYear = m2 > m1 || (m2 === m1 && d2 >= d1);
        const adjYear = useSameYear ? y2 : y2 - 1;
        const adjDay = Math.min(d1, daysInMonth(adjYear, m1));
        return e - dateToSerial(adjYear, m1, adjDay);
      }
      return error('#NUM!');
    },
  },
  {
    name: 'DAYS',
    category: 'datetime',
    minArgs: 2,
    maxArgs: 2,
    signature: [['end_date', 'start_date']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 2) return error('#VALUE!');
      const end = parseSerial(values[0]!, ctx);
      if (isError(end)) return end;
      const start = parseSerial(values[1]!, ctx);
      if (isError(start)) return start;
      return Math.trunc(end) - Math.trunc(start);
    },
  },
  {
    name: 'DAYS360',
    category: 'datetime',
    minArgs: 2,
    maxArgs: 3,
    signature: [['start_date', 'end_date', '[method]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const sVal = ctx.scalar(ctx.evaluate(args[0]!));
      const eVal = ctx.scalar(ctx.evaluate(args[1]!));
      const start = parseSerial(sVal, ctx);
      if (isError(start)) return start;
      const end = parseSerial(eVal, ctx);
      if (isError(end)) return end;
      let euro = false;
      if (args[2] && !isOmitted(ctx.evaluate(args[2]))) {
        const mScalar = ctx.scalar(ctx.evaluate(args[2]));
        if (isError(mScalar)) return mScalar;
        const b = logicalValue(mScalar);
        if (b === undefined) return error('#VALUE!');
        euro = b;
      }
      let [y1, m1, d1] = excelDateParts(Math.trunc(start));
      let [y2, m2, d2] = excelDateParts(Math.trunc(end));
      if (euro) {
        d1 = Math.min(d1, 30);
        d2 = Math.min(d2, 30);
      } else {
        if (d1 === 31 || (m1 === 2 && d1 === daysInMonth(y1, 2))) d1 = 30;
        if (d2 === 31 && d1 === 30) d2 = 30;
      }
      return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
    },
  },
  {
    name: 'WEEKDAY',
    category: 'datetime',
    minArgs: 1,
    maxArgs: 2,
    signature: [['serial_number', '[return_type]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 2) return error('#VALUE!');
      const sVal = ctx.scalar(ctx.evaluate(args[0]!));
      const serial = parseSerial(sVal, ctx);
      if (isError(serial)) return serial;
      let rt = 1;
      if (args[1] && !isOmitted(ctx.evaluate(args[1]))) {
        const rtScalar = ctx.scalar(ctx.evaluate(args[1]));
        if (isError(rtScalar)) return rtScalar;
        rt = Math.trunc(ctx.number(rtScalar));
        if (!Number.isFinite(rt)) return error('#VALUE!');
      }
      return weekdayNumber(serial, rt);
    },
  },
  {
    name: 'WEEKNUM',
    category: 'datetime',
    minArgs: 1,
    maxArgs: 2,
    signature: [['serial_number', '[return_type]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 2) return error('#VALUE!');
      const sVal = ctx.scalar(ctx.evaluate(args[0]!));
      const serial = parseSerial(sVal, ctx);
      if (isError(serial)) return serial;
      let rt = 1;
      if (args[1] && !isOmitted(ctx.evaluate(args[1]))) {
        const rtScalar = ctx.scalar(ctx.evaluate(args[1]));
        if (isError(rtScalar)) return rtScalar;
        rt = Math.trunc(ctx.number(rtScalar));
        if (!Number.isFinite(rt)) return error('#VALUE!');
      }
      if (rt === 21) return isoWeekInfo(serial);
      const startDow = rt === 1 || rt === 17 ? 0 : rt === 2 || rt === 11 ? 1 : rt >= 12 && rt <= 16 ? rt - 10 : -1;
      if (startDow < 0) return error('#NUM!');
      const s = Math.trunc(serial);
      const [y] = excelDateParts(s);
      const jan1 = dateToSerial(y, 1, 1);
      const jan1Offset = (excelDow(jan1) - startDow + 7) % 7;
      return Math.floor((s - jan1 + jan1Offset) / 7) + 1;
    },
  },
  {
    name: 'ISOWEEKNUM',
    category: 'datetime',
    minArgs: 1,
    maxArgs: 1,
    signature: [['date']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const serial = parseSerial(values[0]!, ctx);
      if (isError(serial)) return serial;
      return isoWeekInfo(serial);
    },
  },
  {
    name: 'WORKDAY',
    category: 'datetime',
    minArgs: 2,
    maxArgs: 3,
    signature: [['start_date', 'days', '[holidays]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const start = parseSerial(ctx.scalar(ctx.evaluate(args[0]!)), ctx);
      if (isError(start)) return start;
      const daysVal = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(daysVal)) return daysVal;
      const count = Math.trunc(ctx.number(daysVal));
      if (!Number.isFinite(count)) return error('#VALUE!');
      const holidays = parseHolidays(args[2], ctx);
      if (isError(holidays)) return holidays;
      let cur = Math.trunc(start);
      const step = count >= 0 ? 1 : -1;
      let rem = Math.abs(count);
      while (rem > 0) {
        cur += step;
        const dow = excelDow(cur);
        if (dow !== 0 && dow !== 6 && !holidays.has(cur)) rem--;
      }
      return cur >= 0 && cur <= lastSerial ? cur : error('#NUM!');
    },
  },
  {
    name: 'WORKDAY.INTL',
    category: 'datetime',
    minArgs: 2,
    maxArgs: 4,
    signature: [['start_date', 'days', '[weekend]', '[holidays]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 4) return error('#VALUE!');
      const start = parseSerial(ctx.scalar(ctx.evaluate(args[0]!)), ctx);
      if (isError(start)) return start;
      const daysVal = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(daysVal)) return daysVal;
      const count = Math.trunc(ctx.number(daysVal));
      if (!Number.isFinite(count)) return error('#VALUE!');
      const mask = parseWeekendMask(args[2], ctx, true);
      if (isError(mask)) return mask;
      const holidays = parseHolidays(args[3], ctx);
      if (isError(holidays)) return holidays;
      let cur = Math.trunc(start);
      const step = count >= 0 ? 1 : -1;
      let rem = Math.abs(count);
      while (rem > 0) {
        cur += step;
        if (!mask[excelDow(cur)] && !holidays.has(cur)) rem--;
      }
      return cur >= 0 && cur <= lastSerial ? cur : error('#NUM!');
    },
  },
  {
    name: 'NETWORKDAYS',
    category: 'datetime',
    minArgs: 2,
    maxArgs: 3,
    signature: [['start_date', 'end_date', '[holidays]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const s = parseSerial(ctx.scalar(ctx.evaluate(args[0]!)), ctx);
      if (isError(s)) return s;
      const e = parseSerial(ctx.scalar(ctx.evaluate(args[1]!)), ctx);
      if (isError(e)) return e;
      const holidays = parseHolidays(args[2], ctx);
      if (isError(holidays)) return holidays;
      const mask = [true, false, false, false, false, false, true];
      return countNetworkDays(Math.trunc(s), Math.trunc(e), mask, holidays);
    },
  },
  {
    name: 'NETWORKDAYS.INTL',
    category: 'datetime',
    minArgs: 2,
    maxArgs: 4,
    signature: [['start_date', 'end_date', '[weekend]', '[holidays]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 4) return error('#VALUE!');
      const s = parseSerial(ctx.scalar(ctx.evaluate(args[0]!)), ctx);
      if (isError(s)) return s;
      const e = parseSerial(ctx.scalar(ctx.evaluate(args[1]!)), ctx);
      if (isError(e)) return e;
      const mask = parseWeekendMask(args[2], ctx);
      if (isError(mask)) return mask;
      const holidays = parseHolidays(args[3], ctx);
      if (isError(holidays)) return holidays;
      return countNetworkDays(Math.trunc(s), Math.trunc(e), mask, holidays);
    },
  },
  {
    name: 'YEARFRAC',
    category: 'datetime',
    minArgs: 2,
    maxArgs: 3,
    signature: [['start_date', 'end_date', '[basis]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const sRaw = parseSerial(ctx.scalar(ctx.evaluate(args[0]!)), ctx);
      if (isError(sRaw)) return sRaw;
      const eRaw = parseSerial(ctx.scalar(ctx.evaluate(args[1]!)), ctx);
      if (isError(eRaw)) return eRaw;
      let basis = 0;
      if (args[2] && !isOmitted(ctx.evaluate(args[2]))) {
        const bVal = ctx.scalar(ctx.evaluate(args[2]));
        if (isError(bVal)) return bVal;
        basis = Math.trunc(ctx.number(bVal));
        if (!Number.isFinite(basis)) return error('#VALUE!');
      }
      if (basis < 0 || basis > 4) return error('#NUM!');
      let s = Math.trunc(sRaw);
      let e = Math.trunc(eRaw);
      if (s === e) return 0;
      if (s > e) [s, e] = [e, s];

      if (basis === 2) return finite((e - s) / 360);
      if (basis === 3) return finite((e - s) / 365);

      let [y1, m1, d1] = excelDateParts(s);
      let [y2, m2, d2] = excelDateParts(e);
      if (basis === 0) {
        if (d1 === 31 || (m1 === 2 && d1 === daysInMonth(y1, 2))) d1 = 30;
        if (d2 === 31 && d1 === 30) d2 = 30;
        return finite(((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) / 360);
      }
      if (basis === 4) {
        d1 = Math.min(d1, 30);
        d2 = Math.min(d2, 30);
        return finite(((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) / 360);
      }
      // Basis 1: Actual/Actual
      const actualDays = e - s;
      if (y1 === y2) return finite(actualDays / (isLeapYear(y1) ? 366 : 365));
      if (y2 === y1 + 1 && (m2 < m1 || (m2 === m1 && d2 <= d1))) {
        const feb29InSpan = (isLeapYear(y1) && (m1 < 2 || (m1 === 2 && d1 <= 29))) ||
          (isLeapYear(y2) && (m2 > 2 || (m2 === 2 && d2 === 29)));
        return finite(actualDays / (feb29InSpan ? 366 : 365));
      }
      let yearDaysSum = 0;
      const yearCount = y2 - y1 + 1;
      for (let y = y1; y <= y2; y++) yearDaysSum += isLeapYear(y) ? 366 : 365;
      return finite(actualDays / (yearDaysSum / yearCount));
    },
  },
];

function countNetworkDays(start: number, end: number, mask: readonly boolean[], holidays: ReadonlySet<number>): number {
  const sign = start <= end ? 1 : -1;
  const lo = Math.min(start, end);
  const hi = Math.max(start, end);
  const workdaysPerWeek = mask.filter(off => !off).length;
  const totalDays = hi - lo + 1;
  const fullWeeks = Math.floor(totalDays / 7);
  const rem = totalDays % 7;
  let total = fullWeeks * workdaysPerWeek;
  const startDow = excelDow(lo);
  for (let i = 0; i < rem; i++) {
    if (!mask[(startDow + i) % 7]) total++;
  }
  for (const h of holidays) {
    if (h >= lo && h <= hi && !mask[excelDow(h)]) total--;
  }
  return sign * total;
}
