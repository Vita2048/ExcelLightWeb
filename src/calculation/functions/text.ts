import { formatNumber } from '../../core/formats/cell-format';
import type { FormulaExpression } from '../../core/formula-references';
import { documentLimits } from '../../core/workbook';
import { parseDateText, parseTimeText } from '../formula-dates';
import {
  arrayValues, error, finite, isArray, isError, isOmitted, isRange, numberValue,
  type FormulaError, type FormulaValue, type Scalar,
} from '../formula-values';
import { evaluateScalarArgs, logicalValue } from './helpers';
import type { FunctionContext, FunctionSignature, FunctionSpec } from './types';

const edgeText: FunctionSignature = [['text', '[num_chars]']];
const singleText: FunctionSignature = [['text']];

// CP1252 mapping for CHAR(128..159)
const cp1252Extra: Readonly<Record<number, number>> = {
  128: 0x20ac, 130: 0x201a, 131: 0x0192, 132: 0x201e, 133: 0x2026, 134: 0x2020, 135: 0x2021, 136: 0x02c6,
  137: 0x2030, 138: 0x0160, 139: 0x2039, 140: 0x0152, 142: 0x017d, 145: 0x2018, 146: 0x2019, 147: 0x201c,
  148: 0x201d, 149: 0x2022, 150: 0x2013, 151: 0x2014, 152: 0x02dc, 153: 0x2122, 154: 0x0161, 155: 0x203a,
  156: 0x0153, 158: 0x017e, 159: 0x0178,
};
const reverseCp1252 = new Map<number, number>(Object.entries(cp1252Extra).map(([code, cp]) => [cp, Number(code)]));

function singleTextSpec(name: 'LEN' | 'TRIM' | 'UPPER' | 'LOWER' | 'CLEAN' | 'PROPER', transform: (text: string) => Scalar): FunctionSpec {
  return {
    name,
    category: 'text',
    minArgs: 1,
    maxArgs: 1,
    signature: singleText,
    evaluate(args, ctx) {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const first = values[0] ?? null;
      if (isError(first)) return first;
      return transform(ctx.text(first));
    },
  };
}

function edgeTextSpec(name: 'LEFT' | 'RIGHT'): FunctionSpec {
  return {
    name,
    category: 'text',
    minArgs: 1,
    maxArgs: 2,
    signature: edgeText,
    evaluate(args, ctx) {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length < 1 || values.length > 2) return error('#VALUE!');
      const first = values[0] ?? null;
      if (isError(first)) return first;
      const text = ctx.text(first);
      const size = values.length === 1 ? 1 : Math.trunc(ctx.number(values[1]!));
      if (!Number.isFinite(size) || size < 0) return error('#VALUE!');
      return name === 'LEFT' ? text.slice(0, size) : text.slice(Math.max(0, text.length - size));
    },
  };
}

function checkTextLength(value: string): Scalar {
  return value.length <= documentLimits.textLength ? value : error('#VALUE!');
}

function parseExcelScalarText(raw: string, ctx: FunctionContext): number {
  const trimmed = raw.trim();
  if (!trimmed) return NaN;
  const pct = trimmed.endsWith('%');
  const body = pct ? trimmed.slice(0, -1).trim() : trimmed;
  const unwrapped = body.startsWith('(') && body.endsWith(')') ? '-' + body.slice(1, -1).trim() : body;
  const stripped = unwrapped.replace(/[$€£¥]/gu, '').trim();
  const numeric = ctx.number(stripped);
  if (Number.isFinite(numeric)) return pct ? numeric / 100 : numeric;
  if (!pct) {
    const dt = parseDateText(trimmed, ctx.locale);
    if (Number.isFinite(dt)) return dt;
    const tm = parseTimeText(trimmed);
    if (Number.isFinite(tm)) return tm;
    const space = trimmed.indexOf(' ');
    if (space > 0) {
      const dPart = parseDateText(trimmed.slice(0, space), ctx.locale);
      const tPart = parseTimeText(trimmed.slice(space + 1));
      if (Number.isFinite(dPart) && Number.isFinite(tPart)) return dPart + tPart;
    }
  }
  return NaN;
}

function searchWildcard(pattern: string, within: string, startZeroBased: number): number {
  const tokens: { kind: 'lit' | 'one' | 'star'; ch?: string }[] = [];
  const pat = pattern.toUpperCase();
  for (let i = 0; i < pat.length; i++) {
    const ch = pat[i]!;
    if (ch === '~' && (pat[i + 1] === '~' || pat[i + 1] === '?' || pat[i + 1] === '*')) {
      tokens.push({ kind: 'lit', ch: pat[++i]! });
    } else if (ch === '?') tokens.push({ kind: 'one' });
    else if (ch === '*') {
      if (tokens.at(-1)?.kind !== 'star') tokens.push({ kind: 'star' });
    } else tokens.push({ kind: 'lit', ch });
  }
  const hay = within.toUpperCase();
  const matchFrom = (pos: number): boolean => {
    let i = pos, j = 0, star = -1, retry = pos;
    while (j < tokens.length) {
      const tok = tokens[j]!;
      if (tok.kind === 'star') {
        star = j++;
        retry = i;
      } else if (i < hay.length && (tok.kind === 'one' || tok.ch === hay[i])) {
        i++;
        j++;
      } else if (star >= 0 && retry < hay.length) {
        j = star + 1;
        i = ++retry;
      } else {
        return false;
      }
    }
    return true;
  };
  for (let start = startZeroBased; start <= hay.length; start++) {
    if (matchFrom(start)) return start + 1;
  }
  return -1;
}

function intOrDefault(expr: FormulaExpression | undefined, fallback: number, ctx: FunctionContext): number | FormulaError {
  if (!expr) return fallback;
  const val = ctx.evaluate(expr);
  if (isOmitted(val)) return fallback;
  const s = ctx.scalar(val);
  if (isError(s)) return s;
  const n = ctx.number(s);
  return Number.isFinite(n) ? Math.trunc(n) : error('#VALUE!');
}

function boolOrDefault(expr: FormulaExpression | undefined, fallback: boolean, ctx: FunctionContext): boolean | FormulaError {
  if (!expr) return fallback;
  const val = ctx.evaluate(expr);
  if (isOmitted(val)) return fallback;
  const s = ctx.scalar(val);
  if (isError(s)) return s;
  const b = logicalValue(s);
  return b === undefined ? error('#VALUE!') : b;
}

function formatStrictScalar(value: Scalar, ctx: FunctionContext): string {
  if (isError(value)) return value.value;
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') return ctx.text(value);
  return `"${value.replaceAll('"', '""')}"`;
}

function splitBeforeAfter(before: boolean, args: readonly FormulaExpression[], ctx: FunctionContext): FormulaValue {
  if (args.length < 2 || args.length > 6) return error('#VALUE!');
  const textScalar = ctx.scalar(ctx.evaluate(args[0]!));
  if (isError(textScalar)) return textScalar;
  const source = ctx.text(textScalar);

  const delimVal = ctx.evaluate(args[1]!, true);
  if (isError(delimVal)) return delimVal;
  const delims: string[] = [];
  if (isRange(delimVal)) {
    for (const r of delimVal) for (const v of ctx.values(r)) {
      if (isError(v)) return v;
      if (v !== null) delims.push(ctx.text(v));
    }
  } else if (isArray(delimVal)) {
    for (const v of arrayValues(delimVal)) {
      if (isError(v)) return v;
      if (v !== null) delims.push(ctx.text(v));
    }
  } else {
    const s = ctx.scalar(delimVal);
    if (isError(s)) return s;
    delims.push(ctx.text(s));
  }
  if (!delims.length) return error('#VALUE!');

  const instance = intOrDefault(args[2], 1, ctx);
  if (isError(instance)) return instance;
  if (instance === 0) return error('#VALUE!');

  const matchMode = intOrDefault(args[3], 0, ctx);
  if (isError(matchMode)) return matchMode;
  if (matchMode !== 0 && matchMode !== 1) return error('#VALUE!');

  const matchEnd = intOrDefault(args[4], 0, ctx);
  if (isError(matchEnd)) return matchEnd;
  if (matchEnd !== 0 && matchEnd !== 1) return error('#VALUE!');

  const notFound = (): FormulaValue => {
    if (args[5]) {
      const fb = ctx.evaluate(args[5], true);
      if (!isOmitted(fb)) return fb;
    }
    return error('#N/A');
  };

  if (delims.some(d => d === '')) {
    if (delims.length === 1) {
      return before ? (instance > 0 ? '' : source) : (instance > 0 ? source : '');
    }
  }

  const hay = matchMode === 1 ? source.toUpperCase() : source;
  const needles = (matchMode === 1 ? delims.map(d => d.toUpperCase()) : delims).filter(d => d.length > 0);
  if (!needles.length) return notFound();
  needles.sort((a, b) => b.length - a.length);

  const hits: { start: number; end: number }[] = [];
  for (let i = 0; i <= source.length;) {
    let matchedLen = 0;
    for (const n of needles) {
      if (hay.startsWith(n, i)) {
        matchedLen = n.length;
        break;
      }
    }
    if (matchedLen > 0) {
      hits.push({ start: i, end: i + matchedLen });
      i += matchedLen;
    } else {
      i++;
    }
  }

  if (matchEnd === 1) {
    if (instance > 0) hits.push({ start: source.length, end: source.length });
    else hits.unshift({ start: 0, end: 0 });
  }

  const idx = instance > 0 ? instance - 1 : hits.length + instance;
  if (idx < 0 || idx >= hits.length) return notFound();
  const hit = hits[idx]!;
  return before ? source.slice(0, hit.start) : source.slice(hit.end);
}

export const textFunctions: readonly FunctionSpec[] = [
  edgeTextSpec('LEFT'),
  edgeTextSpec('RIGHT'),
  {
    name: 'MID',
    category: 'text',
    minArgs: 3,
    maxArgs: 3,
    signature: [['text', 'start_num', 'num_chars']],
    evaluate(args, ctx) {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 3) return error('#VALUE!');
      const first = values[0] ?? null;
      if (isError(first)) return first;
      const text = ctx.text(first);
      const start = Math.trunc(ctx.number(values[1]!)), size = Math.trunc(ctx.number(values[2]!));
      return start >= 1 && size >= 0 && Number.isFinite(start + size) ? text.slice(start - 1, start - 1 + size) : error('#VALUE!');
    },
  },
  singleTextSpec('LEN', text => text.length),
  singleTextSpec('TRIM', text => text.replace(/^ +| +$/gu, '').replace(/ +/gu, ' ')),
  singleTextSpec('UPPER', text => text.toUpperCase()),
  singleTextSpec('LOWER', text => text.toLowerCase()),
  singleTextSpec('CLEAN', text => text.replace(/[\u0000-\u001f]/gu, '')),
  singleTextSpec('PROPER', text =>
    text.replace(/\p{L}+/gu, word => word[0]!.toUpperCase() + word.slice(1).toLowerCase())),
  {
    name: 'CONCAT',
    category: 'text',
    minArgs: 1,
    maxArgs: 255,
    signature: [['text1', '[text2]'], true],
    evaluate(args, ctx): Scalar {
      let out = '';
      for (const arg of args) {
        const val = ctx.evaluate(arg, true);
        if (isError(val)) return val;
        if (isRange(val)) {
          for (const r of val) for (const item of ctx.values(r)) {
            if (isError(item)) return item;
            if (item !== null) out += ctx.text(item);
          }
        } else if (isArray(val)) {
          for (const item of arrayValues(val)) {
            if (isError(item)) return item;
            if (item !== null) out += ctx.text(item);
          }
        } else {
          const s = ctx.scalar(val);
          if (isError(s)) return s;
          if (s !== null) out += ctx.text(s);
        }
      }
      return checkTextLength(out);
    },
  },
  {
    name: 'CONCATENATE',
    category: 'text',
    minArgs: 1,
    maxArgs: 255,
    signature: [['text1', '[text2]'], true],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx);
      if (!Array.isArray(values)) return values;
      let out = '';
      for (const v of values) out += isError(v) ? '' : ctx.text(v);
      return checkTextLength(out);
    },
  },
  {
    name: 'TEXTJOIN',
    category: 'text',
    minArgs: 3,
    maxArgs: 255,
    signature: [['delimiter', 'ignore_empty', 'text1', '[text2]'], true],
    evaluate(args, ctx): Scalar {
      if (args.length < 3) return error('#VALUE!');
      const delimScalar = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(delimScalar)) return delimScalar;
      const delimiter = ctx.text(delimScalar);
      const ignoreScalar = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(ignoreScalar)) return ignoreScalar;
      const ignoreEmpty = logicalValue(ignoreScalar);
      if (ignoreEmpty === undefined) return error('#VALUE!');

      const parts: string[] = [];
      const push = (item: Scalar): FormulaError | undefined => {
        if (isError(item)) return item;
        const s = item === null ? '' : ctx.text(item);
        if (ignoreEmpty && s === '') return undefined;
        parts.push(s);
        return undefined;
      };

      for (const arg of args.slice(2)) {
        const val = ctx.evaluate(arg, true);
        if (isError(val)) return val;
        if (isRange(val)) {
          for (const r of val) {
            const area = r;
            const total = (area.bottom - area.top + 1) * (area.right - area.left + 1);
            if (!ignoreEmpty && total > 10_000) return error('#VALUE!');
            if (!ignoreEmpty) {
              for (let row = area.top; row <= area.bottom; row++) {
                for (let col = area.left; col <= area.right; col++) {
                  const err = push(ctx.read(area, row, col));
                  if (err) return err;
                }
              }
            } else {
              for (const item of ctx.values(area)) {
                const err = push(item);
                if (err) return err;
              }
            }
          }
        } else if (isArray(val)) {
          for (const item of arrayValues(val)) {
            const err = push(item);
            if (err) return err;
          }
        } else {
          const err = push(ctx.scalar(val));
          if (err) return err;
        }
      }
      return checkTextLength(parts.join(delimiter));
    },
  },
  {
    name: 'TEXT',
    category: 'text',
    minArgs: 2,
    maxArgs: 2,
    signature: [['value', 'format_text']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 2) return error('#VALUE!');
      const raw = values[0] ?? null;
      const fmt = ctx.text(values[1]!);
      if (typeof raw === 'number') return formatNumber(raw, fmt, ctx.locale).text;
      if (raw === null) return formatNumber(0, fmt, ctx.locale).text;
      if (typeof raw === 'boolean') return ctx.text(raw);
      const parsed = parseExcelScalarText(raw, ctx);
      if (Number.isFinite(parsed)) return formatNumber(parsed, fmt, ctx.locale).text;
      if (fmt === '@' || fmt.toUpperCase() === 'GENERAL') return raw;
      if (fmt.includes('@')) return fmt.replaceAll('@', raw);
      return raw;
    },
  },
  {
    name: 'VALUE',
    category: 'text',
    minArgs: 1,
    maxArgs: 1,
    signature: singleText,
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const raw = values[0] ?? null;
      if (raw === null) return 0;
      if (typeof raw === 'number') return raw;
      if (typeof raw === 'boolean') return error('#VALUE!');
      const parsed = parseExcelScalarText(raw, ctx);
      return Number.isFinite(parsed) ? finite(parsed) : error('#VALUE!');
    },
  },
  {
    name: 'NUMBERVALUE',
    category: 'text',
    minArgs: 1,
    maxArgs: 3,
    signature: [['text', '[decimal_separator]', '[group_separator]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 3) return error('#VALUE!');
      const s0 = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(s0)) return s0;
      if (s0 === null) return 0;
      if (typeof s0 === 'boolean') return error('#VALUE!');
      if (typeof s0 === 'number') return s0;

      let decimalSep = ctx.text(1.5).replace(/\d/g, '')[0] ?? '.';
      if (args[1]) {
        const dVal = ctx.evaluate(args[1]);
        if (!isOmitted(dVal)) {
          const ds = ctx.scalar(dVal);
          if (isError(ds)) return ds;
          const dt = ctx.text(ds);
          if (!dt.length) return error('#VALUE!');
          decimalSep = dt[0]!;
        }
      }
      let groupSep = decimalSep === ',' ? '.' : ',';
      if (args[2]) {
        const gVal = ctx.evaluate(args[2]);
        if (!isOmitted(gVal)) {
          const gs = ctx.scalar(gVal);
          if (isError(gs)) return gs;
          const gt = ctx.text(gs);
          if (!gt.length) return error('#VALUE!');
          groupSep = gt[0]!;
        }
      }
      if (decimalSep === groupSep) return error('#VALUE!');

      let cleaned = s0.replace(/\s+/gu, '');
      let pctCount = 0;
      while (cleaned.endsWith('%')) {
        pctCount++;
        cleaned = cleaned.slice(0, -1);
      }
      if (!cleaned) return error('#VALUE!');
      const decIndex = cleaned.indexOf(decimalSep);
      if (decIndex >= 0 && cleaned.includes(groupSep, decIndex)) return error('#VALUE!');
      const numeric = numberValue(cleaned, decimalSep, groupSep);
      return Number.isFinite(numeric) ? finite(numeric / 100 ** pctCount) : error('#VALUE!');
    },
  },
  {
    name: 'FIND',
    category: 'text',
    minArgs: 2,
    maxArgs: 3,
    signature: [['find_text', 'within_text', '[start_num]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const findScalar = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(findScalar)) return findScalar;
      const withinScalar = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(withinScalar)) return withinScalar;
      const start = intOrDefault(args[2], 1, ctx);
      if (isError(start)) return start;
      const find = ctx.text(findScalar);
      const within = ctx.text(withinScalar);
      if (start < 1 || start > within.length + 1) return error('#VALUE!');
      const idx = within.indexOf(find, start - 1);
      return idx < 0 ? error('#VALUE!') : idx + 1;
    },
  },
  {
    name: 'SEARCH',
    category: 'text',
    minArgs: 2,
    maxArgs: 3,
    signature: [['find_text', 'within_text', '[start_num]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 2 || args.length > 3) return error('#VALUE!');
      const findScalar = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(findScalar)) return findScalar;
      const withinScalar = ctx.scalar(ctx.evaluate(args[1]!));
      if (isError(withinScalar)) return withinScalar;
      const start = intOrDefault(args[2], 1, ctx);
      if (isError(start)) return start;
      const find = ctx.text(findScalar);
      const within = ctx.text(withinScalar);
      if (start < 1 || start > within.length + 1) return error('#VALUE!');
      const hit = searchWildcard(find, within, start - 1);
      return hit < 0 ? error('#VALUE!') : hit;
    },
  },
  {
    name: 'SUBSTITUTE',
    category: 'text',
    minArgs: 3,
    maxArgs: 4,
    signature: [['text', 'old_text', 'new_text', '[instance_num]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 3 || args.length > 4) return error('#VALUE!');
      const s = evaluateScalarArgs(args.slice(0, 3), ctx, true);
      if (!Array.isArray(s)) return s;
      const source = ctx.text(s[0]!);
      const oldText = ctx.text(s[1]!);
      const newText = ctx.text(s[2]!);
      if (!args[3] || isOmitted(ctx.evaluate(args[3]))) {
        return checkTextLength(oldText === '' ? source : source.replaceAll(oldText, newText));
      }
      const inst = intOrDefault(args[3], 0, ctx);
      if (isError(inst)) return inst;
      if (inst < 1) return error('#VALUE!');
      if (oldText === '') return source;
      let from = 0;
      for (let count = 1; count <= inst; count++) {
        const idx = source.indexOf(oldText, from);
        if (idx < 0) return source;
        if (count === inst) {
          return checkTextLength(source.slice(0, idx) + newText + source.slice(idx + oldText.length));
        }
        from = idx + oldText.length;
      }
      return source;
    },
  },
  {
    name: 'REPLACE',
    category: 'text',
    minArgs: 4,
    maxArgs: 4,
    signature: [['old_text', 'start_num', 'num_chars', 'new_text']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 4) return error('#VALUE!');
      const oldText = ctx.text(values[0]!);
      const start = Math.trunc(ctx.number(values[1]!));
      const count = Math.trunc(ctx.number(values[2]!));
      const newText = ctx.text(values[3]!);
      if (!Number.isFinite(start) || !Number.isFinite(count) || start < 1 || count < 0) return error('#VALUE!');
      return checkTextLength(oldText.slice(0, start - 1) + newText + oldText.slice(start - 1 + count));
    },
  },
  {
    name: 'REPT',
    category: 'text',
    minArgs: 2,
    maxArgs: 2,
    signature: [['text', 'number_times']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 2) return error('#VALUE!');
      const text = ctx.text(values[0]!);
      const count = Math.trunc(ctx.number(values[1]!));
      if (!Number.isFinite(count) || count < 0 || text.length * count > documentLimits.textLength) {
        return error('#VALUE!');
      }
      return text.repeat(count);
    },
  },
  {
    name: 'EXACT',
    category: 'text',
    minArgs: 2,
    maxArgs: 2,
    signature: [['text1', 'text2']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 2) return error('#VALUE!');
      return ctx.text(values[0]!) === ctx.text(values[1]!);
    },
  },
  {
    name: 'CHAR',
    category: 'text',
    minArgs: 1,
    maxArgs: 1,
    signature: [['number']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const n = Math.trunc(ctx.number(values[0]!));
      if (!Number.isFinite(n) || n < 1 || n > 255) return error('#VALUE!');
      return String.fromCharCode(cp1252Extra[n] ?? n);
    },
  },
  {
    name: 'CODE',
    category: 'text',
    minArgs: 1,
    maxArgs: 1,
    signature: singleText,
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const s = ctx.text(values[0]!);
      if (!s.length) return error('#VALUE!');
      const cp = s.charCodeAt(0);
      return reverseCp1252.get(cp) ?? cp;
    },
  },
  {
    name: 'UNICHAR',
    category: 'text',
    minArgs: 1,
    maxArgs: 1,
    signature: [['number']],
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const n = Math.trunc(ctx.number(values[0]!));
      if (!Number.isFinite(n) || n < 1 || n > 0x10ffff) return error('#VALUE!');
      if (n >= 0xd800 && n <= 0xdfff) return error('#N/A');
      return String.fromCodePoint(n);
    },
  },
  {
    name: 'UNICODE',
    category: 'text',
    minArgs: 1,
    maxArgs: 1,
    signature: singleText,
    evaluate(args, ctx): Scalar {
      const values = evaluateScalarArgs(args, ctx, true);
      if (!Array.isArray(values)) return values;
      if (values.length !== 1) return error('#VALUE!');
      const s = ctx.text(values[0]!);
      return s.length ? s.codePointAt(0)! : error('#VALUE!');
    },
  },
  {
    name: 'T',
    category: 'text',
    minArgs: 1,
    maxArgs: 1,
    signature: [['value']],
    evaluate(args, ctx): Scalar {
      if (args.length !== 1) return error('#VALUE!');
      const val = ctx.evaluate(args[0]!, true);
      let s: Scalar;
      if (isRange(val)) {
        if (val.length !== 1) return error('#VALUE!');
        s = ctx.read(val[0]!, val[0]!.top, val[0]!.left);
      } else if (isArray(val)) {
        s = val.at(0, 0);
      } else {
        s = ctx.scalar(val);
      }
      if (isError(s)) return s;
      return typeof s === 'string' ? s : '';
    },
  },
  {
    name: 'TEXTBEFORE',
    category: 'text',
    minArgs: 2,
    maxArgs: 6,
    signature: [['text', 'delimiter', '[instance_num]', '[match_mode]', '[match_end]', '[if_not_found]']],
    evaluate: (args, ctx) => splitBeforeAfter(true, args, ctx),
  },
  {
    name: 'TEXTAFTER',
    category: 'text',
    minArgs: 2,
    maxArgs: 6,
    signature: [['text', 'delimiter', '[instance_num]', '[match_mode]', '[match_end]', '[if_not_found]']],
    evaluate: (args, ctx) => splitBeforeAfter(false, args, ctx),
  },
  {
    name: 'VALUETOTEXT',
    category: 'text',
    minArgs: 1,
    maxArgs: 2,
    signature: [['value', '[format]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 2) return error('#VALUE!');
      const val = ctx.scalar(ctx.evaluate(args[0]!));
      const fmt = intOrDefault(args[1], 0, ctx);
      if (isError(fmt)) return fmt;
      if (fmt !== 0 && fmt !== 1) return error('#VALUE!');
      if (fmt === 1) return formatStrictScalar(val, ctx);
      return isError(val) ? val.value : ctx.text(val);
    },
  },
  {
    name: 'ARRAYTOTEXT',
    category: 'text',
    minArgs: 1,
    maxArgs: 2,
    signature: [['array', '[format]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 2) return error('#VALUE!');
      const val = ctx.evaluate(args[0]!, true);
      if (isError(val)) return val;
      const arr = ctx.arrayOf(val);
      if (!arr) return error('#VALUE!');
      const fmt = intOrDefault(args[1], 0, ctx);
      if (isError(fmt)) return fmt;
      if (fmt !== 0 && fmt !== 1) return error('#VALUE!');
      const render = (item: Scalar) => fmt === 1
        ? formatStrictScalar(item, ctx)
        : (isError(item) ? item.value : ctx.text(item));
      if (fmt === 0) {
        const items: string[] = [];
        for (let r = 0; r < arr.rows; r++) for (let c = 0; c < arr.columns; c++) items.push(render(arr.at(r, c)));
        return checkTextLength(items.join(', '));
      }
      const rows: string[] = [];
      for (let r = 0; r < arr.rows; r++) {
        const cols: string[] = [];
        for (let c = 0; c < arr.columns; c++) cols.push(render(arr.at(r, c)));
        rows.push(cols.join(','));
      }
      return checkTextLength(`{${rows.join(';')}}`);
    },
  },
  {
    name: 'FIXED',
    category: 'text',
    minArgs: 1,
    maxArgs: 3,
    signature: [['number', '[decimals]', '[no_commas]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 3) return error('#VALUE!');
      const numScalar = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(numScalar)) return numScalar;
      const num = ctx.number(numScalar);
      if (!Number.isFinite(num)) return error('#VALUE!');
      const dec = intOrDefault(args[1], 2, ctx);
      if (isError(dec)) return dec;
      const noCommas = boolOrDefault(args[2], false, ctx);
      if (isError(noCommas)) return noCommas;
      if (dec > 127) return error('#VALUE!');

      let rounded = num;
      if (dec < 0) {
        const factor = 10 ** (-dec);
        rounded = Math.round(Math.abs(num) / factor) * factor * Math.sign(num);
      } else {
        rounded = Number(Math.round(Number(Math.abs(num) + 'e+' + dec)) + 'e-' + dec) * Math.sign(num);
      }
      if (!Number.isFinite(rounded)) rounded = num;
      const places = Math.max(0, dec);
      const pattern = noCommas
        ? (places > 0 ? `0.${'0'.repeat(places)}` : '0')
        : (places > 0 ? `#,##0.${'0'.repeat(places)}` : '#,##0');
      return formatNumber(rounded, pattern, ctx.locale).text.replaceAll('\u00a0', ' ');
    },
  },
  {
    name: 'DOLLAR',
    category: 'text',
    minArgs: 1,
    maxArgs: 2,
    signature: [['number', '[decimals]']],
    evaluate(args, ctx): Scalar {
      if (args.length < 1 || args.length > 2) return error('#VALUE!');
      const numScalar = ctx.scalar(ctx.evaluate(args[0]!));
      if (isError(numScalar)) return numScalar;
      const num = ctx.number(numScalar);
      if (!Number.isFinite(num)) return error('#VALUE!');
      const dec = intOrDefault(args[1], 2, ctx);
      if (isError(dec)) return dec;
      if (dec > 127) return error('#VALUE!');

      let rounded = num;
      if (dec < 0) {
        const factor = 10 ** (-dec);
        rounded = Math.round(Math.abs(num) / factor) * factor * Math.sign(num);
      } else {
        rounded = Number(Math.round(Number(Math.abs(num) + 'e+' + dec)) + 'e-' + dec) * Math.sign(num);
      }
      if (!Number.isFinite(rounded)) rounded = num;
      const places = Math.max(0, dec);
      const mag = places > 0 ? `#,##0.${'0'.repeat(places)}` : '#,##0';
      const pattern = ctx.locale.toLowerCase().startsWith('cs')
        ? `${mag}" Kč";-${mag}" Kč"`
        : `$${mag};($${mag})`;
      return formatNumber(rounded, pattern, ctx.locale).text.replaceAll('\u00a0', ' ');
    },
  },
];
