import { describe, expect, it } from 'vitest';
import { calculateNative } from '../../src/calculation/native-calculator';
import { isVolatileFunction } from '../../src/calculation/formula-functions';
import { emptySheet, ownDto, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';

function formula(id: string, row: number, column: number, source: string): CellDocument {
  return { id, row, column, content: { kind: 'formula', source } };
}

function num(id: string, row: number, column: number, value: number): CellDocument {
  return { id, row, column, content: { kind: 'number', value } };
}

function text(id: string, row: number, column: number, value: string): CellDocument {
  return { id, row, column, content: { kind: 'text', value } };
}

function book(main: readonly CellDocument[], extra: { other?: readonly CellDocument[]; names?: WorkbookDocument['names'] } = {}): WorkbookDocument {
  return ownDto({
    id: 'wave-a',
    formatVersion: 1,
    dateSystem: '1900',
    sheets: [
      { ...emptySheet('one', 'Sheet1'), cells: main },
      { ...emptySheet('two', "O'Brien data"), cells: extra.other ?? [] },
    ],
    names: extra.names ?? [],
  });
}

describe('Wave A unit behavior', () => {
  it('uses injected options.random for deterministic RANDARRAY output and marks volatility', () => {
    expect(isVolatileFunction('RANDARRAY')).toBe(true);
    expect(isVolatileFunction('OFFSET')).toBe(true);
    expect(isVolatileFunction('INDIRECT')).toBe(true);
    expect(isVolatileFunction('XLOOKUP')).toBe(false);

    const draws = [0.1, 0.9, 0.25, 0.75];
    let index = 0;
    const doc = book([
      formula('r1', 1, 1, '=RANDARRAY(2,2,10,20,TRUE)'),
      formula('r2', 1, 5, '=SUM(A1#)'),
    ]);
    const result = calculateNative(doc, { random: () => draws[index++ % draws.length]! });
    const r1 = result.updates.find(u => u.target.row === 1 && u.target.column === 1);
    const r2 = result.updates.find(u => u.target.row === 1 && u.target.column === 5);

    expect(r1?.spill?.rows).toBe(2);
    expect(r1?.spill?.columns).toBe(2);
    // Range 10..20 has 11 integers -> floor(0.1*11)=1 (11), floor(0.9*11)=9 (19), floor(0.25*11)=2 (12), floor(0.75*11)=8 (18)
    expect(r1?.spill?.cells.map(c => c?.value)).toEqual([11, 19, 12, 18]);
    expect(r2?.cache?.value).toBe(60);
  });

  it('builds dynamic bounding ranges with general colon operators across INDEX, XLOOKUP, and OFFSET', () => {
    const doc = book([
      num('a1', 1, 1, 1), num('b1', 1, 2, 2), num('c1', 1, 3, 3),
      num('a2', 2, 1, 4), num('b2', 2, 2, 5), num('c2', 2, 3, 6),
      num('a3', 3, 1, 7), num('b3', 3, 2, 8), num('c3', 3, 3, 9),
      formula('idxColon', 1, 5, '=SUM(INDEX(A1:C3,1,1):INDEX(A1:C3,2,2))'),
      formula('xlkColon', 2, 5, '=SUM(A1:XLOOKUP(8,B1:B3,B1:B3))'),
      formula('offColon', 3, 5, '=SUM(OFFSET(A1,1,1):C3)'),
      formula('isRef', 4, 5, '=ISREF(INDEX(A1:C3,1,1):B2)'),
      formula('badColon', 5, 5, '=SUM(42:A1)'),
      formula('errColon', 6, 5, '=INDEX(A1:C3,10,1):A1'),
    ]);
    const res = calculateNative(doc);
    const at = (r: number) => res.updates.find(u => u.target.row === r && u.target.column === 5)?.cache?.value;
    expect(at(1)).toBe(1 + 2 + 4 + 5);
    expect(at(2)).toBe(1 + 2 + 4 + 5 + 7 + 8);
    expect(at(3)).toBe(5 + 6 + 8 + 9);
    expect(at(4)).toBe(true);
    expect(at(5)).toBe('#VALUE!');
    expect(at(6)).toBe('#REF!');
  });

  it('resolves INDIRECT in R1C1 mode, cross-sheet quotes, named ranges, and invalid refs', () => {
    const doc = book(
      [
        num('a1', 1, 1, 15),
        num('b1', 1, 2, 25),
        num('b2', 2, 2, 100),
        formula('rel', 2, 3, '=INDIRECT("R[-1]C[-2]",FALSE)'),
        formula('absR1C1', 3, 3, '=INDIRECT("R1C1:R1C2",FALSE)'),
        formula('crossR1C1', 4, 3, '=INDIRECT("\'O\'\'Brien data\'!R2C1",FALSE)'),
        formula('named', 5, 3, '=SUM(INDIRECT("Prices"))'),
        formula('rowRef', 6, 3, '=SUM(INDIRECT("R1:R1",FALSE))'),
        formula('colRef', 7, 3, '=SUM(INDIRECT("C[-1]",FALSE))'),
        formula('oobR1C1', 8, 3, '=INDIRECT("R[-10]C1",FALSE)'),
        formula('zeroR1C1', 9, 3, '=INDIRECT("R0C1",FALSE)'),
      ],
      {
        other: [text('o1', 1, 1, 'hello'), num('o2', 2, 1, 77)],
        names: [{
          id: 'name-prices',
          name: 'Prices',
          area: { sheetId: 'one', row: 1, column: 1, rows: 1, columns: 2 },
        }],
      },
    );
    const res = calculateNative(doc);
    const rowVal = (r: number) => res.updates.find(u => u.target.row === r && u.target.column === 3);
    expect(rowVal(2)?.cache?.value).toBe(15);
    expect(rowVal(3)?.spill?.cells.map(c => c?.value)).toEqual([15, 25]);
    expect(rowVal(4)?.cache?.value).toBe(77);
    expect(rowVal(5)?.cache?.value).toBe(40);
    expect(rowVal(6)?.cache?.value).toBe(40);
    expect(rowVal(7)?.cache?.value).toBe(125); // C2 in col 7 is col 2 (25 + 100)
    expect(rowVal(8)?.cache?.value).toBe('#REF!');
    expect(rowVal(9)?.cache?.value).toBe('#REF!');
  });

  it('preserves reference metadata through IF, CHOOSE, IFS, and SWITCH into FORMULATEXT and ROW/COLUMN', () => {
    const doc = book([
      num('a1', 1, 1, 10),
      formula('a2', 2, 1, '=A1*3'),
      formula('fText', 1, 3, '=FORMULATEXT(CHOOSE(2,A1,A2))'),
      formula('ifRow', 2, 3, '=ROW(IF(TRUE,A2,A1))'),
      formula('ifsCol', 3, 3, '=COLUMN(IFS(FALSE,A1,TRUE,C1))'),
      formula('swRef', 4, 3, '=ISREF(SWITCH(2,1,A1,2,A2))'),
    ]);
    const res = calculateNative(doc);
    const at = (r: number) => res.updates.find(u => u.target.row === r && u.target.column === 3)?.cache?.value;
    expect(at(1)).toBe('=A1*3');
    expect(at(2)).toBe(2);
    expect(at(3)).toBe(3);
    expect(at(4)).toBe(true);
  });
});
