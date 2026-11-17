import { expect, it } from 'vitest';
import { copyFormula, evaluateProbe, insertFormulaRow } from '../../src/calculation/formula-probe';
import { excelDateParts, importExcelSerial } from '../../src/core/date-serial-probe';
import prototype from '../fixtures/formulas/Expected-values.json';

// Production NativeCalculator assertions in desktop S08Checks.cs and S12Checks.cs.
it.each([
  ['=IF(TRUE,42,1/0)', 42], ['=IF(FALSE,1/0,7)', 7],
  ['=IFERROR(1/0,19)', 19], ['=IFERROR(3,1/0)', 3],
  ['=IFERROR(#REF!,"missing")', 'missing'], ['=ABS(" -7 ")', 7],
  ['=ABS(TRUE)', 1], ['=ABS(FALSE)', 0], ['=ABS("bad")', { error: '#VALUE!' }],
  ['=ABS(1/0)', { error: '#DIV/0!' }], ['=ROUND(2.45,1)', 2.5],
  ['=ROUND(-2.5,0)', -3], ['=ROUND(125,-1)', 130],
  ['=ROUND(125,-2)', 100], ['=ROUND(10.22,2)', 10.22],
  ['=ROUND("3.1415",2)', { error: '#VALUE!' }],
  ['=ROUNDUP(-2.1,0)', -3], ['=ROUNDDOWN(-2.9,0)', -2],
  ['=UNSUPPORTED_FUNCTION(1)', { error: '#NAME?' }], ['=SUM(', { error: '#VALUE!' }],
])('evaluates %s with the desktop value or error', (formula, expected) => {
  expect(evaluateProbe(formula)).toEqual(expected);
});

it('never reads a reference in an inactive IF or IFERROR branch', () => {
  const lookup = (address: string) => {
    if (address !== "'Other sheet'!A1") throw new Error(`Unexpected lookup ${address}`);
    return 7;
  };
  expect(evaluateProbe("=IF(TRUE,IFERROR('Other sheet'!$A$1,B2),C3)", lookup)).toBe(7);
  expect(evaluateProbe("='A$ sheet'!$A$1", address => address)).toBe("'A$ sheet'!A1");
});

it('rewrites parsed references without changing quoted text or absolute copy coordinates', () => {
  const expected = prototype.cases.find(c => c.id === 'relative-copy')!;
  expect(copyFormula(expected.expression, expected.origin!, expected.destination!)).toBe(expected.translated);
  expect(copyFormula('=IF(TRUE,\'Other sheet\'!A1,"A1")', 'D1', 'E2')).toBe('=IF(TRUE,\'Other sheet\'!B2,"A1")');
  expect(insertFormulaRow('=SUM(B1:B2)+$C$1', 1)).toBe('=SUM(B2:B3)+$C$2');
  expect(copyFormula('=A1', 'B1', 'A1')).toBe('=#REF!');
});

it('preserves the 1900 leap-day serial and converts only 1904 date values', () => {
  expect([59, 60, 60.5, 61].map(serial => excelDateParts(serial))).toEqual([
    [1900, 2, 28], [1900, 2, 29], [1900, 2, 29], [1900, 3, 1],
  ]);
  expect(excelDateParts(0, 1904)).toEqual([1904, 1, 1]);
  expect(importExcelSerial(1.5, 1904, 'date')).toBe(1463.5);
  expect(importExcelSerial(0.5, 1904, 'time')).toBe(0.5);
  expect(importExcelSerial(10, 1904, 'number')).toBe(10);
  expect(importExcelSerial(60, 1900, 'date')).toBe(60);
});
