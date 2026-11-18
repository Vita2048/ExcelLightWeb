import { describe, expect, it } from 'vitest';
import fixture from '../fixtures/formulas/S08-desktop.json';
import { calculateNative } from '../../src/calculation/native-calculator';
import { nativeFunctions } from '../../src/calculation/formula-functions';
import { emptySheet, ownDto, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';

function formula(id: string, row: number, column: number, source: string): CellDocument {
  return { id, row, column, content: { kind: 'formula', source } };
}

function cells(prefix: string, data: (number | boolean | string)[][]): CellDocument[] {
  return data.map(([row, column, value], i) => {
    const id = prefix + i;
    if (typeof value === 'string' && value.startsWith('=')) return formula(id, Number(row), Number(column), value);
    return { id, row: Number(row), column: Number(column), content:
      typeof value === 'number' ? { kind: 'number', value } :
      typeof value === 'boolean' ? { kind: 'boolean', value } : { kind: 'text', value: String(value) } };
  });
}
function workbook(main: readonly CellDocument[], other: readonly CellDocument[] = [], filtered = false): WorkbookDocument {
  return ownDto({
    id: 'oracle', formatVersion: 1, dateSystem: '1900',
    sheets: [{ ...emptySheet('one', 'Sheet1'), cells: main, ...(filtered && {
      rows: [{ index: 4, hidden: true }],
      filter: { row: 1, column: 1, rows: 4, columns: 2, criteria: [{ column: 1, allowed: ['a'] }] },
    }) }, { ...emptySheet('two', "O'Brien data"), cells: other }],
    names: [],
  });
}

for (const input of fixture) {
  describe(input.id, () => {
    const main = cells('main', input.cells);
    const other = cells('other', input.other);
    it.each(input.cases)('$expression matches the desktop calculator', ({ expression, kind, value }) => {
      const document = workbook([...main, formula('test', 1, 20, expression)], other, input.id === 'subtotal-hidden');
      const result = calculateNative(document, { locale: 'cs-CZ' });
      const cache = result.updates.find(update => update.target.column === 20)?.cache;
      expect(cache?.kind).toBe(kind);
      if (kind === 'number') expect(cache?.value).toBeCloseTo(Number(value), 11);
      else if (kind === 'boolean') expect(cache?.value).toBe(value === 'true');
      else expect(cache?.value).toBe(value);
    });
  });
}

it('exercises valid and invalid arguments for all 31 functions and evaluates TODAY from a deterministic clock', () => {
  const cases = fixture.flatMap(input => input.cases);
  for (const kind of ['result', 'error']) {
    const examples = cases.filter(test => (test.kind === 'error') === (kind === 'error'));
    const names = new Set(examples.flatMap(test =>
      [...test.expression.matchAll(/([A-Z]+)\(/giu)].map(match => match[1]!.toUpperCase())));
    expect(nativeFunctions.filter(name => !names.has(name)), kind).toEqual([]);
  }
  expect(nativeFunctions).toHaveLength(31);
  let clockReads = 0;
  const result = calculateNative(workbook([formula('today', 1, 1, '=TODAY()'), formula('next', 1, 2, '=TODAY()+A1')]), {
    clock: () => { clockReads++; return new Date(2024, 0, 15, 23, 59); },
  });
  expect(clockReads).toBe(1);
  expect(result.updates.map(update => update.cache?.value)).toEqual([45306, 90612]);
});

it('tracks formulas inside the sum range that SUMIF expands from a single cell', () => {
  const document = workbook([
    formula('sum', 1, 1, '=SUMIF(IF(TRUE,C1:C3,D1),1,B1)'),
    formula('value', 3, 2, '=17'),
    { id: 'match', row: 3, column: 3, content: { kind: 'number', value: 1 } },
  ]);
  const result = calculateNative(document);
  expect(result.updates.find(update => update.target.column === 1)?.cache?.value).toBe(17);
  expect(result.dependencies.get('sum')).toEqual([{ sheetId: 'one', row: 3, column: 2 }]);
});

it('evaluates long operator chains within the parser limits', () => {
  const operators = '+1'.repeat(3000);
  const result = calculateNative(workbook([
    formula('chain', 1, 1, '=1' + operators),
    formula('array-chain', 1, 2, '=SUM({1,2}' + operators + ')'),
  ]));
  expect(result.updates.map(update => update.cache?.value)).toEqual([3001, 6003]);
});

it('traverses inactive references and reports complete cycles while preserving lazy IF results', () => {
  const document = workbook([
    formula('a', 1, 1, '=B1+C1'), formula('b', 1, 2, '=A1'), formula('c', 1, 3, '=B1'),
    formula('dependent', 1, 4, '=A1*2'), formula('inactive', 1, 5, '=IF(TRUE,42,E1)'),
    formula('lazy', 1, 6, '=IF(TRUE,42,D1)'), formula('missing', 1, 7, '=IFERROR(A1,"missing")'),
  ]);
  const result = calculateNative(document);
  expect([...result.diagnostics.keys()].sort()).toEqual(['a', 'b', 'c', 'inactive']);
  const byColumn = new Map(result.updates.map(({ target, cache }) => [target.column, cache?.value]));
  expect([...Array(5)].map((_, i) => byColumn.get(i + 1))).toEqual(Array(5).fill('#REF!'));
  expect(byColumn.get(6)).toBe(42);
  expect(byColumn.get(7)).toBe('missing');
  expect(result.dependencies.get('lazy')).toEqual([{ sheetId: 'one', row: 1, column: 4 }]);
});

it('evaluates a 500-cell reverse-order chain and preserves the source and external cache', () => {
  const links = Array.from({ length: 500 }, (_, i) => formula('link-' + i, i + 1, 1, i === 499 ? '=1' : '=A' + (i + 2) + '+1'));
  const external: CellDocument = { ...formula('external', 1, 2, '=NEWFUNCTION(1)'), requiresExternalCalculation: true,
    cache: { kind: 'number', value: 33 } };
  const document = workbook([...links, external, formula('dependent', 1, 3, '=B1+1')]);
  const { updates } = calculateNative(document);
  const result = (column: number) => updates.find(update => update.target.column === column && update.target.row === 1)?.cache;
  expect(result(1)?.value).toBe(500);
  expect(result(2)).toEqual(external.cache);
  expect(result(3)?.value).toBe('#NAME?');
  expect(document.sheets[0]?.cells[500]?.content).toEqual(external.content);
});
