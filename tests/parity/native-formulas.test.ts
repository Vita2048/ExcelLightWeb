import { describe, expect, it } from 'vitest';
import fixture from '../fixtures/formulas/S08-desktop.json';
import smokeExcel from '../fixtures/formulas/smoke-excel.json';
import dynamicArrays from '../fixtures/formulas/dynamic-arrays.json';
import logicalInfo from '../fixtures/formulas/logical-info.json';
import lookupOracle from '../fixtures/formulas/lookup.json';
import shapingOracle from '../fixtures/formulas/shaping.json';
import mathOracle from '../fixtures/formulas/math.json';
import textOracle from '../fixtures/formulas/text.json';
import datetimeOracle from '../fixtures/formulas/datetime.json';
import statsFinancialOracle from '../fixtures/formulas/stats-financial.json';
import getpivotdataOracle from '../fixtures/formulas/getpivotdata.json';
import { calculateNative } from '../../src/calculation/native-calculator';
import { nativeFunctions, supportedFunctions } from '../../src/calculation/formula-functions';
import { createPivotTable, parsePivotRangeReference, parsePivotTargetCell, updatePivotLayout } from '../../src/application/pivot-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import type { PivotAggregation } from '../../src/core/pivot';
import { emptySheet, ownDto, type CellCache, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';

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

function assertOracleScalar(cache: CellCache | undefined, kind: string, value: string, digits = 11) {
  expect(cache?.kind).toBe(kind);
  if (kind === 'number') expect(cache?.value).toBeCloseTo(Number(value), digits);
  else if (kind === 'boolean') expect(cache?.value).toBe(value === 'true');
  else expect(cache?.value).toBe(value);
}

interface OracleCase {
  readonly expression: string;
  readonly kind: string;
  readonly value?: string;
  readonly rows?: number;
  readonly columns?: number;
  readonly matrix?: readonly (readonly { readonly kind: string; readonly value: string }[])[];
}

interface OracleSuite {
  readonly id: string;
  readonly excelVersion: string;
  readonly excelBuild: string;
  readonly cells: unknown;
  readonly other: unknown;
  readonly cases: readonly OracleCase[];
}

function runOracleSuite(suite: OracleSuite, locale: string, skipReason?: (entry: OracleCase) => string | undefined) {
  describe(suite.id, () => {
    const main = cells(`${suite.id}-main`, suite.cells as (number | boolean | string)[][]);
    const other = cells(`${suite.id}-other`, suite.other as (number | boolean | string)[][]);

    for (const entry of suite.cases) {
      const reason = skipReason?.(entry);
      if (reason) {
        it.skip(`${entry.expression} matches Excel oracle (${reason})`, () => {});
        continue;
      }
      it(`${entry.expression} matches Excel ${suite.excelVersion} build ${suite.excelBuild}`, () => {
        const document = workbook([...main, formula('test', 1, 20, entry.expression)], other);
        const result = calculateNative(document, { locale });
        const update = result.updates.find(u => u.target.column === 20);
        if (entry.kind === 'array' && entry.matrix) {
          const rows = entry.rows ?? entry.matrix.length;
          const columns = entry.columns ?? entry.matrix[0]?.length ?? 0;
          expect(update?.spill).toBeDefined();
          expect(update?.spill?.rows).toBe(rows);
          expect(update?.spill?.columns).toBe(columns);
          for (let r = 0; r < rows; r++) {
            for (let c = 0; c < columns; c++) {
              const expected = entry.matrix[r]![c]!;
              const actual = update!.spill!.cells[r * columns + c] ?? undefined;
              assertOracleScalar(actual, expected.kind, expected.value);
            }
          }
        } else {
          expect(update?.spill).toBeUndefined();
          const digits = /\bXIRR\(/iu.test(entry.expression) ? 8 : 11;
          assertOracleScalar(update?.cache, entry.kind, entry.value ?? '', digits);
        }
      });
    }
  });
}

const usesUnsupported = (expression: string) =>
  [...expression.matchAll(/([A-Z.]+)\s*\(/giu)].some(match => !supportedFunctions.has(match[1]!.toUpperCase()));

runOracleSuite(smokeExcel as OracleSuite, 'cs-CZ', entry => {
  if (entry.expression.includes('SEQUENCE(2)')) return 'requires setup blocker outside smoke fixture';
  if (usesUnsupported(entry.expression)) return 'pending function sprint';
  return undefined;
});
runOracleSuite(dynamicArrays as OracleSuite, 'en-US');
runOracleSuite(logicalInfo as OracleSuite, 'en-US');
runOracleSuite(lookupOracle as OracleSuite, 'en-US');
runOracleSuite(shapingOracle as OracleSuite, 'en-US');
runOracleSuite(mathOracle as OracleSuite, 'en-US');
runOracleSuite(textOracle as OracleSuite, 'cs-CZ');
runOracleSuite(datetimeOracle as OracleSuite, 'en-US');
runOracleSuite(statsFinancialOracle as OracleSuite, 'en-US');

it('exercises valid and invalid arguments for all 226 functions and evaluates TODAY from a deterministic clock', () => {
  const cases = [
    ...fixture.flatMap(input => input.cases),
    ...dynamicArrays.cases,
    ...logicalInfo.cases,
    ...lookupOracle.cases,
    ...shapingOracle.cases,
    ...mathOracle.cases,
    ...textOracle.cases,
    ...datetimeOracle.cases,
    ...statsFinancialOracle.cases,
    ...getpivotdataOracle.cases,
  ];
  for (const kind of ['result', 'error']) {
    const examples = cases.filter(test => (test.kind === 'error') === (kind === 'error'));
    const names = new Set(examples.flatMap(test =>
      [...test.expression.matchAll(/([A-Z][A-Z0-9.]*)\(/giu)].map(match => match[1]!.toUpperCase())));
    expect(nativeFunctions.filter(name => !names.has(name)), kind).toEqual([]);
  }
  expect(nativeFunctions).toHaveLength(226);
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

it('sums and looks up a sparse full column, including a blank gap', () => {
  const started = performance.now();
  const document = workbook([
    { id: 'a1', row: 1, column: 1, content: { kind: 'number', value: 10 } },
    { id: 'b1', row: 1, column: 2, content: { kind: 'number', value: 1 } },
    { id: 'a5', row: 5, column: 1, content: { kind: 'number', value: 20 } },
    { id: 'b5', row: 5, column: 2, content: { kind: 'number', value: 2 } },
    formula('sum', 1, 3, '=SUM(A:A)'),
    formula('exact', 2, 3, '=VLOOKUP(20,A:B,2,FALSE)'),
    formula('gap', 3, 3, '=VLOOKUP(15,A:B,2,TRUE)'),
    formula('hit', 4, 3, '=VLOOKUP(20,A:B,2,TRUE)'),
  ]);
  const result = calculateNative(document);
  expect(performance.now() - started).toBeLessThan(50);
  const value = (row: number) => result.updates.find(update => update.target.row === row && update.target.column === 3)?.cache?.value;
  expect(value(1)).toBe(30);
  expect(value(2)).toBe(2);
  expect(value(3)).toBe(1);
  expect(value(4)).toBe(2);
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

describe(getpivotdataOracle.id, () => {
  let idSeq = 0;
  const makeId = () => `gpd-${++idSeq}`;
  const main = cells('gpd-main', getpivotdataOracle.cells as (number | boolean | string)[][]);
  const session = new WorkbookSession(workbook(main), { makeId });
  for (const spec of getpivotdataOracle.pivots) {
    const hiddenByField = new Map<string, string[]>(
      spec.hiddenItems.map(rule => [rule.field, [...rule.items]]),
    );
    const source = parsePivotRangeReference(session.active, 'one', spec.source)!;
    const target = parsePivotTargetCell(session.active, 'one', spec.target)!;
    session.execute(createPivotTable(session.capture(), {
      name: spec.name,
      source,
      target,
      rows: spec.rows.map(fieldName => ({
        fieldName,
        ...(hiddenByField.has(fieldName) ? { hiddenItems: hiddenByField.get(fieldName)! } : {}),
      })),
      columns: spec.columns.map(fieldName => ({
        fieldName,
        ...(hiddenByField.has(fieldName) ? { hiddenItems: hiddenByField.get(fieldName)! } : {}),
      })),
      filters: spec.filters.map(fieldName => ({
        fieldName,
        ...(hiddenByField.has(fieldName) ? { hiddenItems: hiddenByField.get(fieldName)! } : {}),
      })),
      values: spec.values.map(v => ({
        id: makeId(),
        fieldName: v.field,
        aggregation: v.aggregation as PivotAggregation,
        customName: v.customName,
      })),
      grandTotals: spec.grandTotals,
    }, makeId));
    if (spec.collapsed.length > 0) {
      const createdId = session.active.pivots![session.active.pivots!.length - 1]!.id;
      session.execute(updatePivotLayout(session.capture(), createdId, {
        collapsedKeys: spec.collapsed.map(item => item.item),
      }, makeId));
    }
  }
  const pivotDoc = session.active;
  const baseCells = pivotDoc.sheets[0]!.cells;

  it.each(getpivotdataOracle.cases)('$expression matches real Excel (16.0 build 17932)', ({ expression, kind, value }) => {
    const docWithFormula: WorkbookDocument = ownDto({
      ...pivotDoc,
      sheets: pivotDoc.sheets.map((s, idx) => idx === 0
        ? { ...s, cells: [...baseCells, formula('test-gpd', 1, 20, expression)] }
        : s),
    });
    const result = calculateNative(docWithFormula, { locale: 'en-US' });
    const update = result.updates.find(u => u.target.column === 20);
    assertOracleScalar(update?.cache, kind, value ?? '');
  });
});
