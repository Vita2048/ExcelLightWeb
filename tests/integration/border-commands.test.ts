import { expect, it } from 'vitest';
import { applyBorderPreset, applyBorderStroke, inspectBorders } from '../../src/application/border-commands';
import { formatStyles } from '../../src/application/style-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { borderPresets, borderStyles, cellBorders } from '../../src/core/formats/borders';
import { emptySheet, findCell, type CellBorder } from '../../src/core/workbook';

const pen = { style: 'dashDot', rgb: '0070C0' };
const red = { style: 'double', rgb: 'FF0000' };

function fixture(merge = false) {
  const base = emptySheet('sheet', 'Sheet');
  let id = 0;
  const session = new WorkbookSession({ id: 'book', formatVersion: 1, dateSystem: '1900', names: [],
    sheets: [{ ...base, merges: merge ? [{ row: 2, column: 2, rows: 2, columns: 3 }] : [] }] }, { makeId: () => 'id-' + ++id });
  const range = { sheetId: base.id, row: 2, column: 2, rows: 2, columns: 3 };
  session.execute({ kind: 'set-cells', edits: [{ target: range, content: { kind: 'formula', source: '=42' },
    style: { bold: true }, preserveStyle: false }] });
  const { kind: _kind, document: _document, ...basis } = session.beginNativeCalculation('border');
  session.tryCommitNativeCalculation({ ...basis, kind: 'native-calculated', updates: [{ target: range, cache: { kind: 'number', value: 42 } }] });
  const sheet = () => session.active.sheets[0]!;
  const cell = (row: number, column: number) => findCell(sheet(), { row, column });
  const edges = (row: number, column: number) => cellBorders(cell(row, column)?.style);
  return { session, range, sheet, cell, edges };
}

it('applies every preset and style in one transaction and preserves formula identity, cache and font', () => {
  const { session, range, cell, edges } = fixture();
  const baseline = session.active;
  for (const preset of Object.keys(borderPresets) as (keyof typeof borderPresets)[]) {
    const result = session.execute(applyBorderPreset(session.capture(), range, preset, pen));
    expect(cell(2, 2)).toMatchObject({ id: baseline.sheets[0]!.cells[0]!.id, content: { source: '=42' },
      cache: { value: 42 }, style: { bold: true } });
    if (result.changed) {
      const after = session.active;
      expect(session.undo()).toBe(true);
      expect(session.active).toEqual(baseline);
      expect(session.redo()).toBe(true);
      expect(session.active).toEqual(after);
      session.undo();
    }
  }
  for (const style of Object.keys(borderStyles).filter(style => style !== 'none')) {
    const ink = { style, rgb: 'FF0000' };
    session.execute(applyBorderPreset(session.capture(), range, 'top', ink));
    expect(edges(2, 2).top).toEqual(ink);
    expect(edges(1, 2).bottom).toEqual(ink);
  }
  session.execute(applyBorderPreset(session.capture(), range, 'topDoubleBottom', null, pen.rgb));
  expect(edges(2, 2).top).toEqual({ ...pen, style: 'thin' });
  expect(edges(3, 2).bottom).toEqual({ ...pen, style: 'double' });
});

it('mirrors strokes and selectively erases both faces and one diagonal in a single undo step', () => {
  const { session, range, edges } = fixture();
  session.execute(applyBorderPreset(session.capture(), range, 'all', pen));
  session.execute(applyBorderStroke(session.capture(), range.sheetId, [{ row: 2, column: 2, edge: 'diagonalDown' }], pen));
  session.execute(applyBorderStroke(session.capture(), range.sheetId, [{ row: 2, column: 2, edge: 'diagonalUp' }], red));
  expect(edges(2, 2).diagonalDown).toEqual(red);
  const before = session.active;
  session.execute(applyBorderStroke(session.capture(), range.sheetId, [
    { row: 2, column: 3, edge: 'left' }, { row: 3, column: 2, edge: 'top' }, { row: 2, column: 2, edge: 'diagonalUp' },
  ], null));
  expect(edges(2, 2)).toEqual({ top: pen, left: pen, diagonalDown: red });
  expect(edges(2, 3).left).toBeUndefined();
  expect(edges(3, 2).top).toBeUndefined();
  expect(session.undo()).toBe(true);
  expect(session.active).toEqual(before);
  expect(() => applyBorderStroke(session.capture(), range.sheetId, [{ edge: 'left', row: 0, column: 1 }], pen)).toThrow(/Invalid border/);
  expect(session.active).toEqual(before);
});

it('inspects legacy and mixed borders, preserves untouched sides, and maps merged outlines to the anchor', () => {
  const { session, range, edges } = fixture();
  session.execute({ kind: 'set-cells', edits: [{ target: { ...range, column: 3 }, content: { kind: 'blank' },
    preserveStyle: false, style: { borders: 't' } }] });
  expect(inspectBorders(session.capture(), range).edges.top).toBe('mixed');
  const before = session.active;
  session.execute(formatStyles(session.capture(), [range], { fontSize: 18 }, '0.0', { bottom: red }));
  expect(edges(2, 2).top).toBeUndefined();
  expect(edges(2, 3).top).toEqual({ style: 'thin' });
  expect(edges(3, 2).bottom).toEqual(red);
  expect(session.undo()).toBe(true);
  expect(session.active).toEqual(before);

  const merged = fixture(true);
  const target = { ...merged.range, rows: 1, columns: 1 };
  merged.session.execute(applyBorderPreset(merged.session.capture(), target, 'all', red));
  expect(merged.edges(2, 2)).toEqual({ top: red, left: red, right: red, bottom: red });
  expect(inspectBorders(merged.session.capture(), target)).toMatchObject({ rows: 1, columns: 1, edges: { bottom: red, right: red } });
  expect(applyBorderStroke(merged.session.capture(), target.sheetId, [{ edge: 'left', row: 2, column: 3 }], null).edits).toEqual([]);
  merged.session.execute(applyBorderStroke(merged.session.capture(), target.sheetId, [{ edge: 'top', row: 4, column: 3 }], null));
  expect(merged.edges(2, 2).bottom).toBeUndefined();
});

it('replaces and clears both faces of a dialog outline while preserving neighboring content and styles', () => {
  const { session, range, edges, cell } = fixture(true);
  const neighbor = { ...range, row: 4, column: 3 };
  session.execute({ kind: 'set-cells', edits: [{ target: neighbor, content: { kind: 'text', value: 'Neighbor' },
    preserveStyle: false, style: { italic: true, borderDetails: { right: pen } } }] });
  session.execute(applyBorderPreset(session.capture(), range, 'outside', pen));
  const before = session.active;

  session.execute(formatStyles(session.capture(), [range], { fontSize: 18 }, '0.00',
    { top: red, bottom: red, left: red, right: red }));
  expect(edges(2, 2)).toEqual({ top: red, bottom: red, left: red, right: red });
  expect(edges(1, 2).bottom).toEqual(red);
  expect(edges(2, 1).right).toEqual(red);
  expect(edges(3, 5).left).toEqual(red);
  expect(cell(4, 3)).toMatchObject({ content: { kind: 'text', value: 'Neighbor' },
    style: { italic: true, borderDetails: { top: red, right: pen } } });
  expect(cell(4, 3)?.style?.fontSize).toBeUndefined();
  expect(cell(4, 3)?.style?.numberFormat).toBeUndefined();
  const changed = session.active;

  session.execute(formatStyles(session.capture(), [range], {}, undefined,
    { top: null, bottom: null, left: null, right: null }));
  expect(edges(2, 2)).toEqual({});
  expect(edges(1, 2)).toEqual({});
  expect(edges(2, 1)).toEqual({});
  expect(edges(3, 5)).toEqual({});
  expect(edges(4, 3)).toEqual({ right: pen });
  expect(session.undo()).toBe(true);
  expect(session.active).toEqual(changed);
  expect(session.undo()).toBe(true);
  expect(session.active).toEqual(before);
});
