import { describe, expect, it } from 'vitest';
import {
  autoSum, captureClipboard, clear, clipboardText, fill, parseEntry, pasteBlock, pasteText, type EditRange,
} from '../../src/application/editing-commands';
import { WorkbookSession, type CellEdit } from '../../src/application/workbook-session';
import { calculateNative } from '../../src/calculation/native-calculator';
import { copyReferences } from '../../src/core/formula-references';
import { emptySheet, findReferencedCell, type WorkbookDocument } from '../../src/core/workbook';

function workbook(document?: WorkbookDocument) {
  let n = 0;
  const a = emptySheet('input', 'Inputs');
  const b = emptySheet('other', "O'Brien");
  const session = new WorkbookSession(document ?? { id: 'book', formatVersion: 1, dateSystem: '1900', sheets: [a, b], names: [] },
    { makeId: () => 'generated-' + ++n });
  const range = (row: number, column: number, rows = 1, columns = 1, sheetId = a.id): EditRange =>
    ({ sheetId, row, column, rows, columns });
  const at = (row: number, column: number, sheetId = a.id) => ({ sheetId, row, column });
  const find = (row: number, column: number, sheetId = a.id) => findReferencedCell(session.active, at(row, column, sheetId));
  const set = (...edits: CellEdit[]) => session.execute({ kind: 'set-cells', edits });
  function calculate() {
    const { document: _document, kind: _kind, ...basis } = session.beginNativeCalculation('test');
    session.tryCommitNativeCalculation({ ...basis, kind: 'native-calculated', updates: calculateNative(session.active).updates });
  }
  return { session, range, at, find, set, calculate, makeId: () => 'cell-' + ++n };
}
const text = (value: string) => ({ kind: 'text', value }) as const;
const number = (value: number) => ({ kind: 'number', value }) as const;
const formula = (source: string) => ({ kind: 'formula', source }) as const;
const blank = { kind: 'blank' } as const;
const style = { bold: true, fillRgb: 'FFEE00' };

it('keeps a pending cut valid through a cache-only commit and rejects it after an authored edit and undo', () => {
  const { session, range, at, set, calculate, find, makeId } = workbook();
  set({ target: at(1, 1), content: formula('=2+3') });
  const basis = session.capture();
  const cut = captureClipboard(basis, range(1, 1), true);
  calculate();
  expect(session.isCurrent(basis)).toBe(false);
  expect(session.isCurrentAuthored(basis)).toBe(true);
  expect(find(1, 1)?.cache).toEqual(number(5));
  expect(() => pasteBlock(session.capture(), range(2, 1), cut, 'all', makeId)).not.toThrow();
  set({ target: at(1, 2), content: number(7) });
  expect(session.undo()).toBe(true);
  expect(session.sourceFingerprint).toBe(basis.sourceFingerprint);
  expect(session.isCurrentAuthored(basis)).toBe(false);
  expect(() => pasteBlock(session.capture(), range(2, 1), cut, 'all', makeId)).toThrow(/Cut source changed/);
});

describe('S06 typed entry and TSV', () => {
  it('parses finite numbers, Booleans, raw formulas, text styles and deferred Python entries', () => {
    expect(parseEntry('')).toEqual(blank);
    expect(parseEntry('  +01.25e2 ')).toEqual(number(125));
    expect(parseEntry(' False ')).toEqual({ kind: 'boolean', value: false });
    expect(parseEntry('1e999')).toEqual(text('1e999'));
    expect(parseEntry(' 007 ', { numberFormat: '@' })).toEqual(text(' 007 '));
    expect(parseEntry('=1+2')).toEqual(formula('=1+2'));
    expect(parseEntry('=PY()', { numberFormat: '@' })).toEqual(text('=PY()'));
    expect(() => parseEntry('=Py()')).toThrow(/identity/);
  });

  it('preserves blank, empty text, leading zeros, quotes and multiline fields through TSV', () => {
    const { session: s, range, find, makeId } = workbook();
    const tsv = '\t""\t"001"\r\n1\t"a\tb"\t"line\n""two"""';
    s.execute(pasteText(s.capture(), range(1, 1), tsv + '\r\n', 'all', 'en-US', makeId));
    expect(find(1, 1)).toBeUndefined();
    expect(find(1, 2)?.content).toEqual(text(''));
    expect(find(1, 3)?.content).toEqual(text('001'));
    expect(find(2, 3)?.content).toEqual(text('line\n"two"'));
    expect(clipboardText(captureClipboard(s.capture(), range(1, 1, 2, 3), false))).toBe(tsv);
    expect(s.undoCount).toBe(1);
    expect(s.undo()).toBe(true);
    expect(s.active.sheets[0]?.cells).toHaveLength(0);
  });

  it('rejects ragged, malformed, out-of-bounds, oversized and invalid-last-cell batches atomically', () => {
    const { session: s, range, makeId } = workbook();
    const before = s.capture();
    for (const [target, value] of [
      [range(1, 1), '1\t2\n3'], [range(1, 1), '"bad'], [range(1, 1), '"x"junk'],
      [range(10000, 100), '1\t2'], [range(1, 1), '1\t='], [range(1, 1), 'x'.repeat(4 * 1024 * 1024 + 1)],
      [range(1, 1, 2, 2), '1\t2'],
    ] as const) {
      expect(() => s.execute(pasteText(before, target, value, 'all', 'en-US', makeId))).toThrow();
      expect(s.capture()).toBe(before);
      expect(s.undoCount).toBe(0);
    }
  });

  it('keeps formula-looking external text inert in Values and skips nonformulas in Formulas', () => {
    const { session: s, range, find, makeId } = workbook();
    s.execute(pasteText(s.capture(), range(1, 1), '=1+2\t"001"', 'values', 'en-US', makeId));
    expect(find(1, 1)?.content).toEqual(text('=1+2'));
    s.execute(pasteText(s.capture(), range(1, 1), '=4\t"=5"', 'formulas', 'en-US', makeId));
    expect(find(1, 1)?.content).toEqual(formula('=4'));
    expect(find(1, 2)?.content).toEqual(text('001'));
  });
});

describe('S06 clipboard transactions and references', () => {
  it('rewrites reference tokens and source spans without changing functions, literals or names', () => {
    expect(copyReferences('=LOG10(A1)+SUM(\'Other sheet\'!A1:A2)+$B1+C$2+$D$3+"A1"+Data_1', 2, 3))
      .toBe('=LOG10(D3)+SUM(\'Other sheet\'!D3:D4)+$B3+F$2+$D$3+"A1"+Data_1');
    expect(copyReferences('=SUM(A1:B2)+A:A+$1:$2', -1, 1)).toBe('=SUM(#REF!)+B:B+$1:$2');
    expect(() => copyReferences('=SUM(', 1, 0)).toThrow(/malformed/);
    expect(() => copyReferences('=SUM(Sheet1:Sheet2!A1)', 1, 0)).toThrow(/unsupported/);
    expect(() => copyReferences("='[Book.xlsx]Sheet1'!A1", 1, 0)).toThrow(/unsupported/);
  });

  it('copies snapshot styles with new identity and supports values and formulas separately', () => {
    const { session: s, range, at, find, set, calculate, makeId } = workbook();
    set({ target: at(1, 1), content: number(2) },
      { target: at(2, 1), content: formula('=A1+$A$1'), style, preserveStyle: false },
      { target: at(2, 2), content: text('kept text') });
    calculate();
    const frozen = captureClipboard(s.capture(), range(2, 1, 1, 2), false);
    const single = captureClipboard(s.capture(), range(2, 1), false);
    const sourceId = find(2, 1)?.id;
    s.execute(pasteBlock(s.capture(), range(3, 2), frozen, 'all', makeId));
    const copy = find(3, 2)!;
    expect(copy).toMatchObject({ content: formula('=B2+$A$1'), style });
    expect(copy.id).not.toBe(sourceId);
    set({ target: at(4, 1), content: text('prior'), style: { italic: true }, preserveStyle: false });
    s.execute(pasteBlock(s.capture(), range(4, 1), single, 'values', makeId));
    expect(find(4, 1)).toMatchObject({ content: number(4), style: { italic: true } });
    s.execute(pasteBlock(s.capture(), range(4, 1), frozen, 'formulas', makeId));
    expect(find(4, 1)).toMatchObject({ content: formula('=A3+$A$1'), style: { italic: true } });
    expect(find(4, 2)).toBeUndefined();
    set({ target: at(4, 2), content: formula('=1/0') });
    calculate();
    const invalid = captureClipboard(s.capture(), range(4, 2), false);
    expect(() => pasteBlock(s.capture(), range(5, 1), invalid, 'values', makeId)).toThrow(/valid scalar/);
  });

  it('moves overlapping cells and dependent absolute references in one undo and evaluates partial range unions', () => {
    const { session: s, range, at, find, set, calculate, makeId } = workbook();
    set(...[1, 2, 3].map((n, i) => ({ target: at(10, i + 1), content: number(n) })),
      { target: at(11, 1), content: formula('=A10+$B$10') },
      { target: at(11, 2), content: formula('=SUM(A10:C10)') });
    calculate();
    const before = s.capture();
    const cut = captureClipboard(before, range(10, 1, 1, 2), true);
    expect(find(10, 1)?.content).toEqual(number(1));
    s.execute(pasteBlock(s.capture(), range(10, 2), cut, 'all', makeId));
    expect(find(10, 1)).toBeUndefined();
    expect(find(10, 2)?.id).toBe(cut.cells[0]?.id);
    expect(find(10, 3)?.id).toBe(cut.cells[1]?.id);
    expect(find(11, 1)?.content).toEqual(formula('=B10+$C$10'));
    expect(find(11, 2)?.content).toEqual(formula('=SUM((C10,B10:C10))'));
    calculate();
    expect(find(11, 1)?.cache).toEqual(number(3));
    expect(find(11, 2)?.cache).toEqual(number(5));
    expect(s.undoCount).toBe(2);
    expect(s.undo()).toBe(true);
    expect(s.active).toEqual(before.document);
  });

  it('transfers cross-sheet relationships and name identities and rejects stale or non-All cuts', () => {
    const { session: s, range, at, find, set, calculate, makeId } = workbook();
    set({ target: at(1, 1), content: number(3) }, { target: at(2, 1), content: formula('=A1*2') },
      { target: at(2, 2, 'other'), content: formula('=Inputs!A2') });
    s.execute({ kind: 'set-cells', edits: [], names: [{ id: 'name', name: 'MovedCell', area: range(2, 1) }] });
    const cut = captureClipboard(s.capture(), range(2, 1), true);
    expect(() => pasteBlock(s.capture(), range(1, 1, 1, 1, 'other'), cut, 'values', makeId)).toThrow(/not All/);
    s.execute(pasteBlock(s.capture(), range(1, 1, 1, 1, 'other'), cut, 'all', makeId));
    calculate();
    expect(find(1, 1, 'other')).toMatchObject({ id: cut.cells[0]?.id, content: formula("='Inputs'!A1*2"), cache: number(6) });
    expect(find(2, 2, 'other')?.content).toEqual(formula("='O''Brien'!A1"));
    expect(s.active.names[0]).toMatchObject({ id: 'name', area: { sheetId: 'other', row: 1 } });
    expect(() => pasteBlock(s.capture(), range(3, 1), cut, 'all', makeId)).toThrow(/Cut source changed/);
  });

  it('copies a whole merge and rejects partial overwrite without modifying the earlier merge', () => {
    const { session: s, range, at, set, find, makeId } = workbook();
    set({ target: at(1, 1), content: number(8) });
    s.execute({ kind: 'set-cells', edits: [], merges: [{ sheetId: 'input', ranges: [{ row: 1, column: 1, rows: 1, columns: 2 }] }] });
    const merged = captureClipboard(s.capture(), range(1, 2), false);
    expect(merged.range).toEqual(range(1, 1, 1, 2));
    s.execute(pasteBlock(s.capture(), range(2, 1), merged, 'all', makeId));
    expect(find(2, 1)?.content).toEqual(number(8));
    expect(s.active.sheets[0]?.merges).toHaveLength(2);
    const before = s.capture();
    expect(() => pasteBlock(before, range(2, 2), merged, 'all', makeId)).toThrow(/complete merged area/);
    expect(s.capture()).toBe(before);
    expect(() => fill(before, range(1, 1, 2, 2), 'down', makeId)).toThrow(/Unmerge/);
  });
});

describe('S06 Fill, Clear and AutoSum', () => {
  it.each([
    ['down', 2, 1, 4, 1, '=D5+$B$2+C$4+$D7'],
    ['up', 4, 1, 2, 1, '=D1+$B$2+C$4+$D3'],
    ['right', 2, 1, 2, 3, '=F3+$B$2+E$4+$D5'],
    ['left', 2, 3, 2, 1, '=B3+$B$2+A$4+$D5'],
  ] as const)('fills %s from the correct edge with one undo and exact redo', (direction, sr, sc, tr, tc, expected) => {
    const { session: s, range, at, find, set, makeId } = workbook();
    const vertical = direction === 'up' || direction === 'down';
    set({ target: at(sr, sc), content: formula('=D3+$B$2+C$4+$D5'), style, preserveStyle: false },
      { target: at(sr + +!vertical, sc + +vertical), content: blank, style, preserveStyle: false },
      { target: at(tr + +!vertical, tc + +vertical), content: number(9) });
    const before = s.capture();
    s.execute(fill(before, range(2, 1, 3, 3), direction, makeId));
    expect(find(tr, tc)).toMatchObject({ content: formula(expected), style });
    expect(find(tr + +!vertical, tc + +vertical)).toMatchObject({ content: blank, style });
    expect(s.undoCount).toBe(2);
    const after = s.active;
    expect(s.undo() && s.active).toEqual(before.document);
    expect(s.redo() && s.active).toEqual(after);
  });

  it('evaluates filled formulas and AutoSum variants and recalculates after undo without adding history', () => {
    const { session: s, range, at, find, set, calculate, makeId } = workbook();
    set({ target: at(1, 1), content: number(4) }, { target: at(2, 1), content: number(6) },
      { target: at(1, 2), content: formula('=A1*2') });
    s.execute(fill(s.capture(), range(1, 2, 2), 'down', makeId));
    calculate();
    expect(find(2, 2)?.cache).toEqual(number(12));
    for (const [aggregate, value] of [['SUM', 20], ['AVERAGE', 10], ['COUNT', 2], ['MIN', 8], ['MAX', 12]] as const) {
      const result = autoSum(s.capture(), range(1, 2, 3), aggregate);
      expect(result.target).toEqual(at(3, 2));
      s.execute(result.command);
      calculate();
      expect(find(3, 2)).toMatchObject({ content: formula('=' + aggregate + '(B1:B2)'), cache: number(value) });
      expect(s.undoCount).toBe(3);
      s.undo();
      calculate();
      expect(find(3, 2)).toBeUndefined();
    }
    expect(autoSum(s.capture(), range(3, 1), 'SUM').command.edits[0]?.content).toEqual(formula('=SUM(A1:A2)'));
    expect(() => autoSum(s.capture(), range(1, 1, 2, 2), 'SUM')).toThrow(/one row or column/);
  });

  it('clears disjoint areas, preserves layout and removes mirrored borders while preserving adjacent content', () => {
    const { session: s, range, at, find, set } = workbook();
    set({ target: at(1, 1), content: number(8), style, preserveStyle: false },
      { target: at(1, 2), content: formula('=A1*2'), style, preserveStyle: false },
      { target: at(2, 2), content: number(2), style: { ...style, borderDetails: { top: { style: 'thin' }, right: { style: 'thick' } } }, preserveStyle: false },
      { target: at(3, 1), content: text('outside') });
    const before = s.capture();
    s.execute(clear(before, [range(1, 1), range(1, 2)], 'contents'));
    expect(find(1, 1)).toMatchObject({ style, content: blank });
    expect(find(1, 2)).toMatchObject({ style, content: blank });
    s.undo();
    s.execute(clear(s.capture(), [range(1, 2)], 'formats'));
    expect(find(1, 2)?.content).toEqual(formula('=A1*2'));
    expect(find(1, 2)?.style).toBeUndefined();
    expect(find(2, 2)).toMatchObject({ content: number(2), style: { borderDetails: { right: { style: 'thick' } } } });
    expect(find(2, 2)?.style?.borderDetails?.top).toBeUndefined();
    s.undo();
    s.execute(clear(s.capture(), [range(1, 2)], 'all'));
    expect(find(1, 2)).toBeUndefined();
    expect(find(3, 1)?.content).toEqual(text('outside'));
    expect(s.undo() && s.active).toEqual(before.document);
  });
});
