import { describe, expect, it } from 'vitest';
import { isVolatileFunction } from '../../src/calculation/formula-functions';
import { calculateNative } from '../../src/calculation/native-calculator';
import { emptySheet, ownDto, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';
import { readProbeParts } from '../../src/persistence/bounded-zip';
import { addXlfnPrefixes, stripXlfnPrefixes } from '../../src/persistence/xlsx/xlsx-formulas';
import { importXlsx, writeXlsx } from '../../src/persistence/xlsx';

function formula(id: string, row: number, column: number, source: string): CellDocument {
  return { id, row, column, content: { kind: 'formula', source } };
}

function num(id: string, row: number, column: number, value: number): CellDocument {
  return { id, row, column, content: { kind: 'number', value } };
}

function text(id: string, row: number, column: number, value: string): CellDocument {
  return { id, row, column, content: { kind: 'text', value } };
}

describe('Wave B formulas, sparse full-column XLOOKUP, and _xlfn. XLSX interchange', () => {
  it('evaluates NOW, RAND, and RANDBETWEEN deterministically via clock and random options', () => {
    expect(isVolatileFunction('NOW')).toBe(true);
    expect(isVolatileFunction('RAND')).toBe(true);
    expect(isVolatileFunction('RANDBETWEEN')).toBe(true);

    let clockReads = 0;
    const draws = [0.25, 0.8];
    let drawIndex = 0;
    const doc = ownDto({
      id: 'wave-b-vol',
      formatVersion: 1,
      dateSystem: '1900',
      sheets: [{
        ...emptySheet('s1', 'Sheet1'),
        cells: [
          formula('today', 1, 1, '=TODAY()'),
          formula('now', 2, 1, '=NOW()'),
          formula('rand', 3, 1, '=RAND()'),
          formula('rb', 4, 1, '=RANDBETWEEN(10,20)'),
        ],
      }],
      names: [],
    } satisfies WorkbookDocument);

    const result = calculateNative(doc, {
      clock: () => {
        clockReads++;
        return new Date(2024, 0, 15, 12, 30, 0);
      },
      random: () => draws[drawIndex++ % draws.length]!,
    });
    expect(clockReads).toBe(1);
    const val = (r: number) => result.updates.find(u => u.target.row === r)?.cache?.value;
    expect(val(1)).toBe(45306);
    expect(val(2)).toBeCloseTo(45306 + (12.5 / 24), 10);
    expect(val(3)).toBe(0.25);
    expect(val(4)).toBe(10 + Math.floor(0.8 * 11));
  });

  it('evaluates full-column XLOOKUP(D5, Products!A:A, Products!C:C, 0) quickly on sparse sheets', () => {
    const doc = ownDto({
      id: 'sparse-xlookup',
      formatVersion: 1,
      dateSystem: '1900',
      sheets: [
        {
          ...emptySheet('sales', 'Sales'),
          cells: [
            num('d5', 5, 4, 203),
            num('g5', 5, 7, 12),
            formula('f5', 5, 6, '=XLOOKUP(D5, Products!A:A, Products!C:C, 0)'),
            formula('h5', 5, 8, '=F5*G5'),
          ],
        },
        {
          ...emptySheet('products', 'Products'),
          cells: [
            num('p1a', 2, 1, 201), text('p1b', 2, 2, 'UltraBook Pro 15'), num('p1c', 2, 3, 1499.99),
            num('p2a', 3, 1, 202), text('p2b', 3, 2, 'Ergonomic Office Chair'), num('p2c', 3, 3, 299.5),
            num('p3a', 4, 1, 203), text('p3b', 4, 2, '4K Monitor 27-inch'), num('p3c', 4, 3, 389),
          ],
        },
      ],
      names: [],
    } satisfies WorkbookDocument);

    const started = performance.now();
    const res = calculateNative(doc);
    expect(performance.now() - started).toBeLessThan(50);
    expect(res.updates.find(u => u.target.column === 6)?.cache?.value).toBe(389);
    expect(res.updates.find(u => u.target.column === 8)?.cache?.value).toBe(389 * 12);
  });

  it('strips and re-applies _xlfn. prefixes on XLSX import/export while ignoring quoted literals', async () => {
    expect(stripXlfnPrefixes('_xlfn.XLOOKUP(D5, Products!A:A, Products!C:C, "_xlfn.keep")'))
      .toBe('XLOOKUP(D5, Products!A:A, Products!C:C, "_xlfn.keep")');
    expect(stripXlfnPrefixes('_xlfn._xlws.FILTER(A1:A5, A1:A5>0)'))
      .toBe('FILTER(A1:A5, A1:A5>0)');
    expect(addXlfnPrefixes('XLOOKUP(D5, Products!A:A, Products!C:C, 0) & "_xlfn.TEXT" & FILTER(A1:A2, B1:B2>0)'))
      .toBe('_xlfn.XLOOKUP(D5, Products!A:A, Products!C:C, 0) & "_xlfn.TEXT" & _xlfn._xlws.FILTER(A1:A2, B1:B2>0)');

    const doc = ownDto({
      id: 'xlfn-roundtrip',
      formatVersion: 1,
      dateSystem: '1900',
      sheets: [
        {
          ...emptySheet('sales', 'Sales'),
          cells: [
            num('d5', 5, 4, 203),
            num('g5', 5, 7, 12),
            formula('f5', 5, 6, '=XLOOKUP(D5,Products!A:A,Products!C:C,0)'),
            formula('h5', 5, 8, '=F5*G5'),
            formula('i5', 5, 9, '=TEXTJOIN("-",TRUE,"A","","B")'),
            formula('j5', 5, 10, '=FILTER(G5:G5,G5:G5>0)'),
          ],
        },
        {
          ...emptySheet('products', 'Products'),
          cells: [
            num('p3a', 4, 1, 203),
            num('p3c', 4, 3, 389),
          ],
        },
      ],
      names: [],
    } satisfies WorkbookDocument);

    const exported = await writeXlsx(doc);
    const parts = await readProbeParts(exported.blob);
    const xmlBytes = parts.get('xl/worksheets/sheet1.xml');
    expect(xmlBytes).toBeDefined();
    const xml = new TextDecoder().decode(xmlBytes);

    expect(xml).toContain('_xlfn.XLOOKUP(');
    expect(xml).toContain('_xlfn.TEXTJOIN(');
    expect(xml).toContain('_xlfn._xlws.FILTER(');

    let nextId = 0;
    const imported = await importXlsx(exported.blob, { makeId: () => `id-${++nextId}` });
    const salesSheet = imported.document.sheets.find(s => s.name === 'Sales')!;
    const f5Cell = salesSheet.cells.find(c => c.row === 5 && c.column === 6)!;
    const h5Cell = salesSheet.cells.find(c => c.row === 5 && c.column === 8)!;
    expect(f5Cell.requiresExternalCalculation).toBeUndefined();
    expect(f5Cell.content).toEqual({ kind: 'formula', source: '=XLOOKUP(D5,Products!A:A,Products!C:C,0)' });
    expect(h5Cell.requiresExternalCalculation).toBeUndefined();

    const calc = calculateNative(imported.document);
    expect(calc.updates.find(u => u.target.row === 5 && u.target.column === 6)?.cache?.value).toBe(389);
    expect(calc.updates.find(u => u.target.row === 5 && u.target.column === 8)?.cache?.value).toBe(389 * 12);
    expect(calc.updates.find(u => u.target.row === 5 && u.target.column === 9)?.cache?.value).toBe('A-B');
    expect(calc.updates.find(u => u.target.row === 5 && u.target.column === 10)?.cache?.value).toBe(12);
  });

  it('formats US TEXT, VALUE, FIXED, DOLLAR, and ARRAYTOTEXT canonically under en-US', () => {
    const doc = ownDto({
      id: 'wave-b-text-us',
      formatVersion: 1,
      dateSystem: '1900',
      sheets: [{
        ...emptySheet('s1', 'Sheet1'),
        cells: [
          formula('t1', 1, 1, '=TEXT(1234.5,"#,##0.00")'),
          formula('t2', 2, 1, '=TEXT(45306,"yyyy-mm-dd")'),
          formula('v1', 3, 1, '=VALUE("12.50")'),
          formula('fx1', 4, 1, '=FIXED(1234.567,1,FALSE)'),
          formula('dl1', 5, 1, '=DOLLAR(1234.567,2)'),
          formula('dl2', 6, 1, '=DOLLAR(-1234.5,0)'),
          formula('at0', 7, 1, '=ARRAYTOTEXT({1,2},0)'),
          formula('at1', 8, 1, '=ARRAYTOTEXT({1,"a";TRUE,2},1)'),
        ],
      }],
      names: [],
    } satisfies WorkbookDocument);
    const res = calculateNative(doc, { locale: 'en-US' });
    const at = (r: number) => res.updates.find(u => u.target.row === r)?.cache?.value;
    expect(at(1)).toBe('1,234.50');
    expect(at(2)).toBe('2024-01-15');
    expect(at(3)).toBe(12.5);
    expect(at(4)).toBe('1,234.6');
    expect(at(5)).toBe('$1,234.57');
    expect(at(6)).toBe('($1,235)');
    expect(at(7)).toBe('1, 2');
    expect(at(8)).toBe('{1,"a";TRUE,2}');
  });
});
