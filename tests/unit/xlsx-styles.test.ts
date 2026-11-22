import { expect, it } from 'vitest';
import { fillPatterns } from '../../src/core/formats/appearance';
import { borderStyles } from '../../src/core/formats/borders';
import type { CellDocument, CellStyle, SheetDocument, WorkbookDocument } from '../../src/core/workbook';
import { importXlsx, writeXlsx } from '../../src/persistence/xlsx';

let sequence = 0;
const makeId = () => 'style-id-' + (++sequence);

function workbookOf(styles: readonly CellStyle[], sheet?: Partial<SheetDocument>): WorkbookDocument {
  const cells: CellDocument[] = styles.map((style, index) => ({
    id: makeId(), row: index + 1, column: 1, content: { kind: 'number', value: index }, style,
  }));
  return {
    id: makeId(), formatVersion: 1, dateSystem: '1900',
    sheets: [{
      id: makeId(), name: 'Styles', cells, rows: [], columns: [], merges: [],
      frozenPanes: { rows: 0, columns: 0 }, ...sheet,
    }],
    names: [],
  };
}

async function roundTripStyles(styles: readonly CellStyle[]): Promise<(CellStyle | undefined)[]> {
  const { blob } = await writeXlsx(workbookOf(styles));
  const { document } = await importXlsx(blob);
  const cells = document.sheets[0]!.cells;
  return styles.map((_, index) => cells.find(cell => cell.row === index + 1)?.style);
}

it('round trips the supported font, number, fill, border and alignment properties', async () => {
  const styles: CellStyle[] = [
    { bold: true, italic: true, strikethrough: true, underline: true },
    { underline: true, doubleUnderline: true },
    { underline: true, underlineKind: 'singleAccounting' },
    { underline: true, underlineKind: 'doubleAccounting', doubleUnderline: true },
    { script: 'superscript', fontFamily: 'Segoe UI', fontSize: 14, textRgb: '1F4E78' },
    { script: 'subscript', fontFamily: 'Arial', fontSize: 72 },
    { numberFormat: '#,##0.00;[Red](#,##0.00)' },
    { numberFormat: '[$-405]dd.mm.yyyy' },
    { fillRgb: '123456' },
    ...Object.keys(fillPatterns).filter(pattern => pattern !== 'none').map(pattern => ({
      fillRgb: 'FFFFFF', fillDetails: { pattern, patternRgb: '000000' },
    })),
    { fillRgb: 'FFFFFF', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'horizontal', variant: 0 } },
    { fillRgb: '112233', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'vertical', variant: 1 } },
    { fillRgb: 'FFFFFF', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'vertical', variant: 2 } },
    { fillRgb: 'FFFFFF', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'diagonalUp', variant: 1 } },
    { fillRgb: 'FFFFFF', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'diagonalDown', variant: 2 } },
    { fillRgb: 'FFFFFF', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'corner', variant: 0 } },
    { fillRgb: 'FFFFFF', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'corner', variant: 1 } },
    { fillRgb: 'FFFFFF', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'corner', variant: 2 } },
    { fillRgb: 'FFFFFF', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'corner', variant: 3 } },
    { fillRgb: 'FFFFFF', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'center', variant: 0 } },
    { fillRgb: 'FFFFFF', fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'center', variant: 2 } },
    { borders: 'trbl' },
    {
      borderDetails: {
        top: { style: 'slantDashDot', rgb: '112233' },
        right: { style: 'double', rgb: '223344' },
        bottom: { style: 'mediumDashDotDot', rgb: '334455' },
        left: { style: 'dashed', rgb: '445566' },
        diagonalUp: { style: 'dotted', rgb: '556677' },
        diagonalDown: { style: 'dotted', rgb: '556677' },
      },
    },
  ];
  expect(Object.keys(borderStyles).filter(style => style !== 'none')).toHaveLength(13);
  await expect(roundTripStyles(styles)).resolves.toEqual(styles);
});

it('maps every border line style, every alignment mode and every text rotation', async () => {
  const edgeStyles = Object.keys(borderStyles).filter(style => style !== 'none');
  const borderCells: CellStyle[] = edgeStyles.map((style, index) => ({
    borderDetails: { top: { style, rgb: String(index).padStart(2, '0').repeat(3) } },
  }));
  const horizontal = ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'] as const;
  const vertical = ['top', 'center', 'justify', 'distributed'] as const;
  const alignmentCells: CellStyle[] = [
    ...horizontal.map(alignment => ({ alignment })),
    ...vertical.map(verticalAlignment => ({ verticalAlignment })),
    { wrap: true, alignment: 'justify' },
    { shrinkToFit: true, verticalAlignment: 'center' },
    { indent: 3, readingOrder: 'ltr' },
    { readingOrder: 'rtl', justifyLastLine: true },
    ...[-90, -45, 45, 90, 255].map(textRotation => ({ textRotation })),
  ];
  const styles = [...borderCells, ...alignmentCells];
  await expect(roundTripStyles(styles)).resolves.toEqual(styles);
});

it('canonicalizes the reversed center gradient order the way the desktop reader does', async () => {
  const original: CellStyle = {
    fillRgb: 'FFFFFF',
    fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'center', variant: 1 },
  };
  const result = await roundTripStyles([original]);
  expect(result[0]).toEqual({
    fillRgb: '2288AA',
    fillDetails: { pattern: 'none', gradientRgb: 'FFFFFF', shading: 'center', variant: 0 },
  });
});

it('preserves merges, explicit dimensions and both freeze axes', async () => {
  const document: WorkbookDocument = {
    id: makeId(), formatVersion: 1, dateSystem: '1900',
    sheets: [{
      id: makeId(), name: 'Geometry',
      cells: [
        { id: makeId(), row: 1, column: 1, content: { kind: 'text', value: 'Merged' } },
        { id: makeId(), row: 4, column: 1, content: { kind: 'number', value: 1 }, style: { bold: true } },
        { id: makeId(), row: 4, column: 3, content: { kind: 'blank' }, style: { italic: true } },
      ],
      rows: [{ index: 2, size: 40 }, { index: 3, size: 26, hidden: true }],
      columns: [{ index: 1, size: 124 }, { index: 3, size: 61, hidden: true }],
      merges: [{ row: 1, column: 1, rows: 2, columns: 2 }],
      frozenPanes: { rows: 1, columns: 2 },
    }],
    names: [],
  };
  const { blob } = await writeXlsx(document);
  const { document: loaded } = await importXlsx(blob);
  const sheet = loaded.sheets[0]!;
  expect(sheet.merges).toEqual([{ row: 1, column: 1, rows: 2, columns: 2 }]);
  expect(sheet.frozenPanes).toEqual({ rows: 1, columns: 2 });
  expect(sheet.rows).toEqual([{ index: 2, size: 40, hidden: false }, { index: 3, size: 26, hidden: true }]);
  expect(sheet.columns).toEqual([{ index: 1, size: 124, hidden: false }, { index: 3, size: 61, hidden: true }]);
});
