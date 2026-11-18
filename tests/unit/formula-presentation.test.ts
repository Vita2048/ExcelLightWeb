import { expect, test } from 'vitest';
import { formulaHint } from '../../src/features/editing/formula-hints';
import { localeArgumentSeparator, presentFormula } from '../../src/features/formulas/formula-presentation';

test('only function argument separators are localized through nested calls and escaped quotes', () => {
  const canonical = `=IF('Sep,;''s'!A1,SUM({1,2;3,4},(A1,B1)),LEFT("a,"";",2))`;
  const displayed = `=IF('Sep,;''s'!A1;SUM({1,2;3,4};(A1,B1));LEFT("a,"";";2))`;
  expect(localeArgumentSeparator('cs-CZ')).toBe(';');
  expect(localeArgumentSeparator('en-US')).toBe(',');
  expect(presentFormula(canonical, ',', ';')).toBe(displayed);
  expect(presentFormula(displayed, ';', ',')).toBe(canonical);
  expect(presentFormula('plain;text', ';', ',')).toBe('plain;text');
  expect(formulaHint(displayed, displayed.lastIndexOf('2'), ';')).toMatchObject({ name: 'LEFT', active: 1 });
});
