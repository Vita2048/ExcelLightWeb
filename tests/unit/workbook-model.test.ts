import { describe, expect, it } from 'vitest';
import { applyDelta, diffWorkbook, invertDelta, sameAuthoredWorkbook } from '../../src/core/workbook-delta';
import { columnLetters } from '../../src/core/formula-references';
import {
  DocumentError, documentLimits, emptySheet, findCell, newWorkbook, ownDto, resolvedCellStyle, worksheetViewport, type CellDocument, type WorkbookDocument,
} from '../../src/core/workbook';
import { validateWorkbook } from '../../src/core/workbook-validation';

const ids = ['book', 'sheet'][Symbol.iterator]();
const book = newWorkbook(() => ids.next().value!);
const sheet = book.sheets[0]!;
const cell = (index: number): CellDocument => ({
  id: `cell-${index}`, row: index % 10_000 + 1, column: Math.floor(index / 10_000) + 1,
  content: { kind: 'number', value: index },
});
const withCells = (cells: readonly CellDocument[]): WorkbookDocument => ({ ...book, sheets: [{ ...sheet, cells }] });

describe('sparse workbook model', () => {
  it('accepts the sheet count and coordinate boundaries and rejects the next sheet', () => {
    const sheets = Array.from({ length: 10 }, (_, index) => emptySheet(`s${index}`, `Sheet ${index}`));
    const edge = { ...cell(0), row: documentLimits.rows, column: documentLimits.columns };
    expect(columnLetters(documentLimits.columns)).toBe('XFD');
    expect(worksheetViewport).toEqual({ rows: documentLimits.rows, columns: documentLimits.columns, address: 'A1:XFD1048576' });
    sheets[0] = { ...sheets[0]!, cells: [edge] };
    expect(() => validateWorkbook({ ...book, sheets })).not.toThrow();
    expect(findCell(sheets[0]!, edge)).toBe(edge);
    expect(findCell(sheets[0]!, { row: 9_999, column: 100 })).toBeUndefined();
    expect(() => validateWorkbook({ ...book, sheets: [...sheets, emptySheet('extra', 'Extra')] })).toThrow(DocumentError);
    expect(() => validateWorkbook({ ...book, sheets: [] })).toThrow(DocumentError);
  });

  it('enforces workbook-wide populated-cell and native formula budgets at the exact boundaries', () => {
    const cells = Array.from({ length: documentLimits.cells + 1 }, (_, index) => cell(index));
    expect(() => validateWorkbook(withCells(cells.slice(0, -1)))).not.toThrow();
    expect(() => validateWorkbook(withCells(cells))).toThrow(/200,000/);
    const formulas = cells.slice(0, documentLimits.formulas + 1).map(entry => ({
      ...entry, content: { kind: 'formula', source: '=1' } as const,
    }));
    expect(() => validateWorkbook(withCells(formulas.slice(0, -1)))).not.toThrow();
    expect(() => validateWorkbook(withCells(formulas))).toThrow(/10,000/);
  });

  it('validates identities across cells, sheets, and names and retains distinct Python anchors', () => {
    const python = { ...cell(0), content: { kind: 'python', source: '1 + 2', blockId: 'block-a' } as const };
    const name = { id: 'range-a', name: 'DataSet', area: { sheetId: sheet.id, row: 1, column: 1, rows: 1, columns: 1 } };
    const document = { ...withCells([python]), names: [name] };
    expect(() => validateWorkbook(document)).not.toThrow();
    expect(() => validateWorkbook({ ...document, names: [{ ...name, id: python.id }] })).toThrow(/identities/);
    expect(() => validateWorkbook(withCells([python, { ...cell(1), content: python.content }]))).toThrow(/block identity/);
    expect(() => validateWorkbook(withCells([python, { ...cell(1), row: python.row }]))).toThrow(/coordinate/);
    expect(() => validateWorkbook({ ...book, sheets: [sheet, emptySheet('other', sheet.name.toLowerCase())] })).toThrow(/sheet name/);
    expect(() => validateWorkbook(withCells([{ ...python, content: { ...python.content, source: 'é'.repeat(524_288) } }]))).not.toThrow();
    expect(() => validateWorkbook(withCells([{ ...python, content: { ...python.content, source: 'é'.repeat(524_289) } }]))).toThrow(/1 MiB/);
  });

  it('validates sheet metadata and rejects covering a populated cell with a merge', () => {
    const state = {
      ...sheet, rows: [{ index: 5, size: 30, hidden: true }], columns: [{ index: 3, size: 50 }],
      filter: { row: 1, column: 3, rows: 2, columns: 2, criteria: [{ column: 4, allowed: ['', 'One'] }] },
      frozenPanes: { rows: 1, columns: 2 }, merges: [{ row: 1, column: 1, rows: 2, columns: 1 }],
      cells: [cell(0)],
    };
    const validate = (changes: Partial<typeof state>) => validateWorkbook({ ...book, sheets: [{ ...state, ...changes }] });
    expect(() => validate({})).not.toThrow();
    expect(() => validate({ cells: [cell(1)] })).toThrow(/merged range/);
    expect(() => validate({ merges: [...state.merges, ...state.merges] })).toThrow(/overlap/);
    expect(() => validate({ rows: [...state.rows, ...state.rows] })).toThrow(/dimension/);
    expect(() => validate({ frozenPanes: { rows: documentLimits.rows, columns: 0 } })).toThrow(/frozen panes/);
    expect(() => validateWorkbook(withCells([{ ...cell(0), row: documentLimits.rows + 1, column: 1 }]))).toThrow(DocumentError);
    expect(() => validateWorkbook(withCells([{ ...cell(0), row: 1, column: documentLimits.columns + 1 }]))).toThrow(DocumentError);
    expect(() => validate({ filter: { ...state.filter, criteria: [{ column: 2, allowed: [] }] } })).toThrow(/filter column/);
  });

  it('accepts 1,000 distinct definitions and rejects excess or overlapping tables', () => {
    const area = { sheetId: sheet.id, row: 1, column: 1, rows: 2, columns: 1 };
    const names = Array.from({ length: 1001 }, (_, i) => ({ id: `range-${i}`, name: `Data_${i}`, area }));
    expect(() => validateWorkbook({ ...book, names: names.slice(0, -1) })).not.toThrow();
    expect(() => validateWorkbook({ ...book, names })).toThrow(/1,000/);
    expect(() => validateWorkbook({ ...book, names: names.slice(0, 2).map(name => ({ ...name, isTable: true })) })).toThrow(/overlap/);
  });

  it('rejects ordinal case duplicates without expanding distinct sheet names', () => {
    const withNames = (first: string, second: string) => ({
      ...book, sheets: [emptySheet('one', first), emptySheet('two', second)],
    });
    expect(() => validateWorkbook(withNames('Čísla', 'ČÍSLA'))).toThrow(/sheet name/);
    expect(() => validateWorkbook(withNames('Straße', 'STRASSE'))).not.toThrow();
  });

  it('accepts exact formula and text lengths and rejects NULs or one extra character', () => {
    const text = 'a'.repeat(documentLimits.textLength);
    const formula = '=' + '1'.repeat(documentLimits.formulaLength - 1);
    const validate = (value: CellDocument['content']) => validateWorkbook(withCells([{ ...cell(0), content: value }]));
    expect(() => validate({ kind: 'formula', source: formula })).not.toThrow();
    expect(() => validate({ kind: 'formula', source: formula + '1' })).toThrow(/formula/);
    expect(() => validate({ kind: 'text', value: text })).not.toThrow();
    expect(() => validate({ kind: 'text', value: text + 'a' })).toThrow(/text/);
    expect(() => validate({ kind: 'text', value: 'a\0' })).toThrow(/NUL/);
  });

  it('ignores sparse and axis enumeration order and derived caches for authored equality', () => {
    const first = ownDto(withCells([cell(0), cell(1)]));
    const second = ownDto(withCells([cell(1), { ...cell(0), cache: { kind: 'error', value: '#REF!' } }]));
    expect(sameAuthoredWorkbook(first, second)).toBe(true);
    expect(sameAuthoredWorkbook(first, withCells([cell(0), { ...cell(1), id: 'new-cell' }]))).toBe(false);
    const sheetWithRows = { ...sheet, rows: [{ index: 4, size: 30 }, { index: 2 }] };
    expect(sameAuthoredWorkbook({ ...book, sheets: [sheetWithRows] },
      { ...book, sheets: [{ ...sheetWithRows, rows: [...sheetWithRows.rows].reverse() }] })).toBe(true);
  });
});

it('round-trips exact deltas over generated mixed edits, additions, deletions and reorders', () => {
  let seed = 0x34e291;
  function random() {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed / 2 ** 32;
  }
  const source = Array.from({ length: 30 }, (_, index) => cell(index));
  for (let trial = 0; trial < 120; trial++) {
    const before = ownDto(withCells(source.filter(() => random() < 0.7)));
    let cells = source.filter(() => random() < 0.7).map(current => random() < 0.5 ? current :
      { ...current, content: { kind: 'text', value: String(current.row) } as const });
    if (random() < 0.5) cells = cells.reverse();
    const after = ownDto(withCells(cells));
    const delta = diffWorkbook(before, after);
    expect(applyDelta(before, delta)).toEqual(after);
    expect(applyDelta(after, invertDelta(delta))).toEqual(before);
  }
});

it('resolves axis styles with cell > row > column precedence, explicit {} overrides, and validates dimension styles', () => {
  const styledSheet = ownDto({
    ...sheet,
    rows: [{ index: 2, style: { italic: true } }],
    columns: [{ index: 3, style: { fillRgb: 'FFFF00' } }],
    cells: [
      { id: 'c-override', row: 3, column: 3, content: { kind: 'blank' as const }, style: {} },
      { id: 'c-explicit', row: 2, column: 3, content: { kind: 'number' as const, value: 1 }, style: { bold: true } },
    ],
  });
  expect(resolvedCellStyle(styledSheet, { row: 1, column: 3 })).toEqual({ fillRgb: 'FFFF00' });
  expect(resolvedCellStyle(styledSheet, { row: 2, column: 1 })).toEqual({ italic: true });
  expect(resolvedCellStyle(styledSheet, { row: 2, column: 3 })).toEqual({ bold: true });
  expect(resolvedCellStyle(styledSheet, { row: 3, column: 3 })).toEqual({});
  expect(resolvedCellStyle(styledSheet, { row: 1, column: 1 })).toBeUndefined();
  expect(() => validateWorkbook({ ...book, sheets: [styledSheet] })).not.toThrow();
  expect(() => validateWorkbook({
    ...book,
    sheets: [{ ...sheet, columns: [{ index: 1, style: { fontSize: 100 } }] }],
  })).toThrow(/6 to 72/);
});
