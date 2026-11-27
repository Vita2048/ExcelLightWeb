import { describe, expect, it } from 'vitest';
import { calculateNative } from '../../src/calculation/native-calculator';
import {
  copyReferences,
  moveReferences,
  parseFormula,
  renameTableColumnReferences,
  renameTableReferences,
  rewriteSheetReferences,
  shiftEnclosedRows,
  shiftReferences,
  UnsupportedFormulaReference,
} from '../../src/core/formula-references';
import { emptySheet, ownDto, type WorkbookDocument } from '../../src/core/workbook';
import { resolveStructuredReference } from '../../src/core/workbook-names';

function salesWorkbook(): WorkbookDocument {
  return ownDto({
    id: 'parser-book',
    formatVersion: 1,
    dateSystem: '1900',
    sheets: [{
      ...emptySheet('s1', 'Sheet1'),
      cells: [
        { id: 'h1', row: 1, column: 1, content: { kind: 'text', value: 'Item' } },
        { id: 'h2', row: 1, column: 2, content: { kind: 'text', value: 'Price' } },
        { id: 'h3', row: 1, column: 3, content: { kind: 'text', value: 'Qty' } },
        { id: 'h4', row: 1, column: 4, content: { kind: 'text', value: 'LineTotal' } },
        { id: 'r2c1', row: 2, column: 1, content: { kind: 'text', value: 'Pen' } },
        { id: 'r2c2', row: 2, column: 2, content: { kind: 'number', value: 10 } },
        { id: 'r2c3', row: 2, column: 3, content: { kind: 'number', value: 3 } },
        { id: 'r2c4', row: 2, column: 4, content: { kind: 'formula', source: '=[@Price]*[@Qty]' } },
        { id: 'r3c1', row: 3, column: 1, content: { kind: 'text', value: 'Book' } },
        { id: 'r3c2', row: 3, column: 2, content: { kind: 'number', value: 25 } },
        { id: 'r3c3', row: 3, column: 3, content: { kind: 'number', value: 2 } },
        { id: 'r3c4', row: 3, column: 4, content: { kind: 'formula', source: '=[@Price]*[@Qty]' } },
        { id: 'tot1', row: 4, column: 1, content: { kind: 'text', value: 'Total' } },
        { id: 'tot4', row: 4, column: 4, content: { kind: 'formula', source: '=SUM(Sales[LineTotal])' } },
        { id: 'ext1', row: 1, column: 6, content: { kind: 'formula', source: '=SUM(Sales[[#Data],[Price]:[Qty]])' } },
        { id: 'ext2', row: 2, column: 6, content: { kind: 'formula', source: '=COUNTA(Sales[#Headers])' } },
        { id: 'ext3', row: 3, column: 6, content: { kind: 'formula', source: '=Sales[[#Totals],[LineTotal]]' } },
        { id: 'ext4', row: 4, column: 6, content: { kind: 'formula', source: '=SUM(Sales[MissingCol])' } },
      ],
    }],
    names: [{
      id: 'sales-tbl',
      name: 'Sales',
      isTable: true,
      area: { sheetId: 's1', row: 1, column: 1, rows: 4, columns: 4 },
      design: { style: 'Blue', headerRow: true, totalRow: true },
    }],
  });
}

describe('Sprint 1.2 formula parser and reference upgrades', () => {
  it('parses spill range (#) and implicit intersection (@) operators', () => {
    expect(parseFormula('=A1#').expression).toMatchObject({
      kind: 'spill',
      target: { kind: 'reference', reference: { first: { row: 1, column: 1 }, range: false } },
    });
    expect(parseFormula('=Sheet2!B2#').expression).toMatchObject({
      kind: 'spill',
      target: { kind: 'reference', reference: { sheet: 'Sheet2', first: { row: 2, column: 2 } } },
    });
    expect(parseFormula('=SUM(A1#)').expression).toMatchObject({
      kind: 'call',
      name: 'SUM',
      args: [{ kind: 'spill', target: { kind: 'reference' } }],
    });

    expect(parseFormula('=@A1:A10').expression).toMatchObject({
      kind: 'unary',
      operator: '@',
      value: { kind: 'reference', reference: { range: true } },
    });
    expect(parseFormula('=@Table1[Col]').expression).toMatchObject({
      kind: 'unary',
      operator: '@',
      value: { kind: 'structured', reference: { table: 'Table1', startColumn: 'Col' } },
    });
    expect(parseFormula('=@A1#').expression).toMatchObject({
      kind: 'unary',
      operator: '@',
      value: { kind: 'spill', target: { kind: 'reference' } },
    });
  });

  it('parses general range (:), intersection (space), and higher-order calls while rejecting malformed numbers', () => {
    expect(parseFormula('=SUM(INDEX(A1:A10,1):INDEX(A1:A10,5))').expression).toMatchObject({
      kind: 'call',
      name: 'SUM',
      args: [{
        kind: 'binary',
        operator: ':',
        left: { kind: 'call', name: 'INDEX' },
        right: { kind: 'call', name: 'INDEX' },
      }],
    });
    expect(parseFormula('=A1:INDEX(A1:A10,5)').expression).toMatchObject({
      kind: 'binary',
      operator: ':',
      left: { kind: 'reference' },
      right: { kind: 'call', name: 'INDEX' },
    });
    expect(parseFormula('=SUM(A1:A4 A2:A3)').expression).toMatchObject({
      kind: 'call',
      name: 'SUM',
      args: [{ kind: 'binary', operator: ' ', left: { kind: 'reference' }, right: { kind: 'reference' } }],
    });
    expect(parseFormula('=SUM((A1:B2) (B1:C2))').expression).toMatchObject({
      kind: 'call',
      name: 'SUM',
      args: [{ kind: 'binary', operator: ' ', left: { kind: 'reference' }, right: { kind: 'reference' } }],
    });
    expect(() => parseFormula('=1 2')).toThrow();

    expect(parseFormula('=LAMBDA(x,x*2)(5)').expression).toMatchObject({
      kind: 'invoke',
      callee: { kind: 'call', name: 'LAMBDA' },
      args: [{ kind: 'literal', value: 5 }],
    });
    expect(parseFormula('=(LAMBDA(x,LAMBDA(y,x+y)))(2)(3)').expression).toMatchObject({
      kind: 'invoke',
      callee: { kind: 'invoke', callee: { kind: 'call', name: 'LAMBDA' }, args: [{ kind: 'literal', value: 2 }] },
      args: [{ kind: 'literal', value: 3 }],
    });
    expect(parseFormula('=LET(f,LAMBDA(x,x+1),f(5))').expression).toMatchObject({
      kind: 'call',
      name: 'LET',
      args: [
        { kind: 'name', name: 'f' },
        { kind: 'call', name: 'LAMBDA' },
        { kind: 'call', name: 'F', args: [{ kind: 'literal', value: 5 }] },
      ],
    });
  });

  it('parses structured table references and resolves/evaluates them against table metadata', () => {
    const cases = [
      ['=Sales[LineTotal]', { table: 'Sales', specifiers: [], startColumn: 'LineTotal' }],
      ['=Sales[[#Headers],[Price]]', { table: 'Sales', specifiers: ['#Headers'], startColumn: 'Price' }],
      ['=Sales[[#Totals],[LineTotal]]', { table: 'Sales', specifiers: ['#Totals'], startColumn: 'LineTotal' }],
      ['=Sales[[#Data],[Price]:[Qty]]', { table: 'Sales', specifiers: ['#Data'], startColumn: 'Price', endColumn: 'Qty' }],
      ['=[@Price]', { specifiers: ['@'], startColumn: 'Price' }],
      ['=Sales[@[Price]]', { table: 'Sales', specifiers: ['@'], startColumn: 'Price' }],
      ["=Sales['[Net']]", { table: 'Sales', specifiers: [], startColumn: '[Net]' }],
    ] as const;

    for (const [source, expected] of cases) {
      const parsed = parseFormula(source);
      expect(parsed.structuredReferences).toHaveLength(1);
      expect(parsed.structuredReferences[0]).toMatchObject(expected);
    }

    const doc = salesWorkbook();
    expect(resolveStructuredReference(doc, parseFormula('=Sales[Price]').structuredReferences[0]!))
      .toEqual({ sheetId: 's1', row: 2, column: 2, rows: 2, columns: 1 });
    expect(resolveStructuredReference(doc, parseFormula('=Sales[[#Headers],[Price]:[Qty]]').structuredReferences[0]!))
      .toEqual({ sheetId: 's1', row: 1, column: 2, rows: 1, columns: 2 });
    expect(resolveStructuredReference(doc, parseFormula('=Sales[[#Totals],[LineTotal]]').structuredReferences[0]!))
      .toEqual({ sheetId: 's1', row: 4, column: 4, rows: 1, columns: 1 });
    expect(resolveStructuredReference(doc, parseFormula('=[@Price]').structuredReferences[0]!, { sheetId: 's1', row: 3, column: 4 }))
      .toEqual({ sheetId: 's1', row: 3, column: 2, rows: 1, columns: 1 });

    const calc = calculateNative(doc);
    const at = (row: number, column: number) =>
      calc.updates.find(u => u.target.row === row && u.target.column === column)?.cache;
    expect(at(2, 4)).toEqual({ kind: 'number', value: 30 });
    expect(at(3, 4)).toEqual({ kind: 'number', value: 50 });
    expect(at(4, 4)).toEqual({ kind: 'number', value: 80 });
    expect(at(1, 6)).toEqual({ kind: 'number', value: 40 });
    expect(at(2, 6)).toEqual({ kind: 'number', value: 4 });
    expect(at(3, 6)).toEqual({ kind: 'number', value: 80 });
    expect(at(4, 6)).toEqual({ kind: 'error', value: '#REF!' });
  });

  it('parses 3D sheet references while rejecting external workbook references', () => {
    const threeD = parseFormula('=SUM(Sheet1:Sheet3!A1:B5)');
    expect(threeD.references).toHaveLength(1);
    expect(threeD.references[0]).toMatchObject({
      sheet: 'Sheet1',
      endSheet: 'Sheet3',
      first: { row: 1, column: 1 },
      last: { row: 5, column: 2 },
      range: true,
    });

    const quotedThreeD = parseFormula("=SUM('Q1 2025:Q4 2025'!$B$2:B10)");
    expect(quotedThreeD.references[0]).toMatchObject({
      sheet: 'Q1 2025',
      endSheet: 'Q4 2025',
      first: { row: 2, column: 2, lockRow: true, lockColumn: true },
      last: { row: 10, column: 2 },
    });

    expect(() => parseFormula('=[1]Sheet1!A1')).toThrow(UnsupportedFormulaReference);
    expect(() => parseFormula('=[Book.xlsx]Sheet1!A1')).toThrow(UnsupportedFormulaReference);
    expect(() => parseFormula("='[Book.xlsx]Sheet1'!A1")).toThrow(/unsupported/);
  });

  it('preserves #, @, structured refs, and 3D refs across structural and table rewrites', () => {
    expect(copyReferences('=SUM(A1#)+@B2:B5+Sheet1:Sheet3!C1', 2, 1))
      .toBe('=SUM(B3#)+@C4:C7+Sheet1:Sheet3!D3');
    expect(shiftEnclosedRows('=SUM(A2#)+@B2', 2, 5, 3)).toBe('=SUM(A5#)+@B5');
    expect(moveReferences('=SUM(A1#)+@B1', 'Sheet1', 'Sheet1', {
      sourceSheet: 'Sheet1', destinationSheet: 'Sheet1',
      from: { top: 1, left: 1, bottom: 1, right: 1 }, dr: 4, dc: 0,
    })).toBe('=SUM(A5#)+@B1');

    const order = ['Q1', 'Q2', 'Q3', 'Q4'];
    expect(rewriteSheetReferences("=SUM(Q1:Q4!A1)+'Q1:Q4'!B2", 'Q1', 'Quarter 1'))
      .toBe("=SUM('Quarter 1:Q4'!A1)+'Quarter 1:Q4'!B2");
    expect(rewriteSheetReferences('=SUM(Q1:Q3!A1)', 'Q1', undefined, order))
      .toBe('=SUM(Q2:Q3!A1)');
    expect(rewriteSheetReferences('=SUM(Q1:Q2!A1)', 'Q1', undefined, order))
      .toBe('=SUM(\'Q2\'!A1)');
    expect(rewriteSheetReferences('=SUM(Q1:Q1!A1)', 'Q1', undefined, order))
      .toBe('=SUM(#REF!)');

    expect(shiftReferences('=SUM(Q1:Q3!A2:A5)', 'Summary', 'Q2', { axis: 'row', start: 2, count: 2, delete: false }, order))
      .toBe('=SUM(Q1:Q3!A4:A7)');
    expect(shiftReferences('=SUM(Q1:Q3!A2:A5)', 'Summary', 'Q4', { axis: 'row', start: 2, count: 2, delete: false }, order))
      .toBe('=SUM(Q1:Q3!A2:A5)');

    expect(renameTableReferences('=SUM(Sales[Amount])+COUNTA(Sales)+"Sales"+SUM(SalesOther[Amount])', 'Sales', 'Orders'))
      .toBe('=SUM(Orders[Amount])+COUNTA(Orders)+"Sales"+SUM(SalesOther[Amount])');

    expect(renameTableColumnReferences('=SUM(Sales[[#Data],[Price]:[Qty]])+[@Price]', 'Sales', 'Price', 'Unit Price'))
      .toBe('=SUM(Sales[[#Data],[Unit Price]:[Qty]])+[@Price]');
    expect(renameTableColumnReferences(
      '=SUM(Sales[[#Data],[Price]:[Qty]])+[@Price]',
      'Sales',
      'Price',
      'Unit [Price]',
      { sheetId: 's1', row: 2, column: 4 },
      { sheetId: 's1', row: 1, column: 1, rows: 4, columns: 4 },
    )).toBe("=SUM(Sales[[#Data],[Unit '[Price']]:[Qty]])+[@Unit '[Price']]");
  });
});
