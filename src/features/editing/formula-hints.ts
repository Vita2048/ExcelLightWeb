import type { NativeFunction } from '../../calculation/formula-functions';

type Signature = readonly [parameters: readonly string[], variadic?: true];
const numbers = [['number1', '[number2]'], true] as const;
const values = [['value1', '[value2]'], true] as const;
const logicals = [['logical1', '[logical2]'], true] as const;
const rounding = [['number', 'num_digits']] as const;
const edgeText = [['text', '[num_chars]']] as const;
const text = [['text']] as const;
const datePart = [['serial_number']] as const;
const signatures: Readonly<Record<NativeFunction, Signature>> = {
  SUM: numbers, AVERAGE: numbers, MIN: numbers, MAX: numbers, COUNT: values, COUNTA: values,
  IF: [['logical_test', '[value_if_true]', '[value_if_false]']], AND: logicals, OR: logicals,
  NOT: [['logical']], IFERROR: [['value', 'value_if_error']], ABS: [['number']],
  ROUND: rounding, ROUNDUP: rounding, ROUNDDOWN: rounding,
  SUMIF: [['range', 'criteria', '[sum_range]']], COUNTIF: [['range', 'criteria']],
  LEFT: edgeText, RIGHT: edgeText, MID: [['text', 'start_num', 'num_chars']],
  LEN: text, TRIM: text, UPPER: text, LOWER: text,
  DATE: [['year', 'month', 'day']], YEAR: datePart, MONTH: datePart, DAY: datePart, TODAY: [[]],
  VLOOKUP: [['lookup_value', 'table_array', 'col_index_num', '[range_lookup]']],
  SUBTOTAL: [['function_num', 'ref1', '[ref2]'], true],
};

export function formulaHint(source: string, caret: number) {
  if (!source.startsWith('=')) return;
  const frames: { name: string; argument: number }[] = [];
  let quote = '', brackets = 0, braces = 0;
  const end = Math.min(caret, source.length);
  for (let i = 1; i < end; i++) {
    const character = source[i]!;
    if (quote) {
      if (character === quote) {
        if (i + 1 < end && source[i + 1] === quote) i++;
        else quote = '';
      }
      continue;
    }
    if (character === '"' || character === "'") { quote = character; continue; }
    if (character === '[') brackets++;
    if (character === ']') brackets = Math.max(0, brackets - 1);
    if (brackets) continue;
    if (character === '{') braces++;
    if (character === '}') braces = Math.max(0, braces - 1);
    if (braces) continue;
    if (character === '(') {
      const name = /[\p{L}\p{N}_.]+$/u.exec(source.slice(1, i).trimEnd())?.[0].toUpperCase() ?? '';
      frames.push({ name, argument: 0 });
    } else if (character === ')') frames.pop();
    else if (character === ',' && frames.at(-1)?.name) frames.at(-1)!.argument++;
  }
  for (const { name, argument } of frames.reverse()) {
    if (Object.hasOwn(signatures, name)) {
      const [parameters, variadic = false] = signatures[name as NativeFunction];
      return { name, parameters, variadic, active: Math.min(argument, parameters.length - 1) };
    }
  }
}
