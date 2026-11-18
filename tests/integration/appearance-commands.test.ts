import { expect, it } from 'vitest';
import { activeFormatBrush, captureFormatBrush, formatStyles, paintStyles } from '../../src/application/style-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { fillPatterns, fillShadings, underlineOptions, underlinePatch } from '../../src/core/formats/appearance';
import { emptySheet, findCell, type CellStyle, type WorkbookDocument } from '../../src/core/workbook';

function fixture() {
  const original: WorkbookDocument = { id: 'book', formatVersion: 1, dateSystem: '1900', names: [],
    sheets: [emptySheet('inputs', 'Inputs'), { ...emptySheet('second', 'Second'),
      merges: [{ row: 3, column: 4, rows: 2, columns: 1 }] }] };
  let next = 0;
  const session = new WorkbookSession(original, { makeId: () => 'generated-' + ++next });
  const range = (row: number, column: number, rows = 1, columns = 1, sheetId = 'inputs') =>
    ({ sheetId, row, column, rows, columns });
  const cell = (row: number, column: number, sheetIndex = 0) => findCell(session.active.sheets[sheetIndex]!, { row, column });
  return { original, session, range, cell };
}

it('preserves all font effects and worksheet fill variants in the model and rejects malformed appearance atomically', () => {
  const { session, range, cell } = fixture();
  const fills = Object.keys(fillPatterns).filter(pattern => pattern !== 'none').map(pattern =>
    ({ pattern, patternRgb: '0070C0' }));
  const gradients = Object.keys(fillShadings).flatMap(shading => [0, 1, 2, 3].map(variant =>
    ({ pattern: 'none', shading, variant, gradientRgb: '5B9BD5' })));
  const effects = [...fills, ...gradients];
  for (const [index, fillDetails] of effects.entries()) {
    const row = index + 1;
    const rangeAt = range(row, 1);
    const before = session.active;
    const style = { fillRgb: 'E7E6E6', fillDetails, fontFamily: 'Custom Worksheet Font', fontSize: index % 2 ? 6 : 72,
      bold: true, italic: true, textRgb: '123456', strikethrough: true,
      script: index % 2 ? 'subscript' : 'superscript',
      ...underlinePatch(underlineOptions[index % underlineOptions.length]!) };
    session.execute(formatStyles(session.capture(), [rangeAt], style));
    expect(cell(row, 1)?.style).toEqual(style);
    const after = session.active;
    expect(session.undo()).toBe(true);
    expect(session.active).toEqual(before);
    expect(session.redo()).toBe(true);
    expect(session.active).toEqual(after);
  }
  const before = session.capture();
  for (const patch of [{ fontSize: 73 }, { script: 'invalid' }, { fillDetails: { pattern: 'invalid' } },
    { fillDetails: { pattern: 'none', shading: 'center' } }, { textRgb: 'abcxyz' }]) {
    expect(() => session.execute(formatStyles(before, [range(1, 1, 2)], patch))).toThrow();
    expect(session.capture()).toBe(before);
  }
  expect(() => formatStyles(before, [range(1, 1)], { bold: true }, '0.00"')).toThrow(/valid Excel/);
  const saved = JSON.parse(JSON.stringify(session.active));
  const malformed = JSON.parse(JSON.stringify(saved));
  malformed.sheets[0].cells[0].style.fillDetails.pattern = ['darkGray'];
  expect(() => session.adopt(malformed)).toThrow(/Invalid fill/);
  expect(session.capture()).toBe(before);
  session.adopt(saved);
  expect(session.active).toEqual(saved);
});

it('combines changed Font, Number and Fill fields in one transaction without flattening mixed styles', () => {
  const { session, range, cell } = fixture();
  session.execute(formatStyles(session.capture(), [range(1, 1)], { bold: true, textRgb: '112233' }, '0.00'));
  session.execute(formatStyles(session.capture(), [range(2, 1)], { italic: true, fillRgb: 'ABCDEF' }, '0.00%'));
  const before = session.active;
  expect(session.execute(formatStyles(session.capture(), [range(1, 1, 2)], {})).changed).toBe(false);
  session.execute(formatStyles(session.capture(), [range(1, 1, 2)], { fontSize: 18, script: 'superscript' }));
  expect(cell(1, 1)?.style).toEqual({ bold: true, textRgb: '112233', numberFormat: '0.00', fontSize: 18, script: 'superscript' });
  expect(cell(2, 1)?.style).toEqual({ italic: true, fillRgb: 'ABCDEF', numberFormat: '0.00%', fontSize: 18, script: 'superscript' });
  expect(session.undo()).toBe(true);
  expect(session.active).toEqual(before);
});

it('paints the detached style pattern across sheets while preserving content, caches, identities and merges', () => {
  const { original, session, range, cell } = fixture();
  const style: CellStyle = { fontFamily: 'Consolas', bold: true, fillRgb: '123456', numberFormat: '0.00',
    wrap: true, alignment: 'right', borders: 'trbl' };
  const target = range(3, 3, 2, 3, 'second');
  session.execute({ kind: 'set-cells', edits: [
    { target: range(1, 1), content: { kind: 'number', value: 12 }, preserveStyle: false, style },
    { target, content: { kind: 'formula', source: '=1+2' } },
    { target: { ...target, column: 4 }, content: { kind: 'text', value: 'keep' }, preserveStyle: false, style: { bold: true } },
  ] });
  const { kind: _kind, document: _document, ...basis } = session.beginNativeCalculation('painter');
  expect(session.tryCommitNativeCalculation({ ...basis, kind: 'native-calculated',
    updates: [{ target, cache: { kind: 'number', value: 3 } }] })).toBe(true);
  const brush = captureFormatBrush(session.capture(), range(1, 1, 1, 2));
  session.execute(formatStyles(session.capture(), [range(1, 1)], { bold: undefined, fillRgb: 'FFEEDD' }));
  const before = session.active;
  const destination = cell(3, 3, 1)!;
  session.execute(paintStyles(session.capture(), target, brush));
  expect(cell(3, 3, 1)).toEqual({ ...destination, style });
  expect(cell(3, 4, 1)).toMatchObject({ content: { kind: 'text', value: 'keep' } });
  expect(cell(3, 4, 1)?.style).toBeUndefined();
  expect(cell(4, 5, 1)).toMatchObject({ content: { kind: 'blank' }, style });
  expect(session.active.sheets[1]?.merges).toEqual(before.sheets[1]?.merges);
  const after = session.active;
  expect(session.undo()).toBe(true);
  expect(session.active).toEqual(before);
  expect(session.redo()).toBe(true);
  expect(session.active).toEqual(after);
  expect(paintStyles(session.capture(), range(6, 1, 1, 1, 'second'), brush).edits).toHaveLength(1);
  expect(() => paintStyles(session.capture(), range(1, 100), brush)).toThrow(/outside/);
  session.adopt(original);
  expect(activeFormatBrush(session.capture(), brush)).toBeUndefined();
  expect(() => paintStyles(session.capture(), target, brush)).toThrow(/different workbook/);
  expect(session.active).toEqual(original);
  expect(session.canUndo).toBe(false);
});
