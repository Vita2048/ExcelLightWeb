import { expect, test } from 'vitest';
import { referenceSlot } from '../../src/features/formulas/reference-pick';

test('recognizes insertion and reference replacement without treating quoted content as a slot', () => {
  for (const marked of ['=|', '=SUM(|)', '=IF(TRUE;|;"text")', '=ROUND(SUM(A1:B2);|)']) {
    const start = marked.indexOf('|');
    expect(referenceSlot(marked.replace('|', ''), { start, end: start })).toEqual({ start, end: start });
  }
  for (const marked of ['|=', '=SUM(|A1)', '="comma,|value"', '="escaped ""quote,|"',
    "=SUM('Sheet;|Name'!A1)", '=SUM({1,|2;3,4})', '=SUM(A1)|']) {
    const start = marked.indexOf('|');
    expect(referenceSlot(marked.replace('|', ''), { start, end: start })).toBeUndefined();
  }
  const existing = "'O''Brien'!\u0024B\u00242:C3";
  const start = 5, end = start + existing.length;
  expect(referenceSlot('=SUM(' + existing + ')', { start, end })).toEqual({ start, end });
  expect(referenceSlot('=SUM("B2")', { start, end: 9 })).toBeUndefined();
});
