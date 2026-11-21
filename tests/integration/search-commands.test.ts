import { describe, expect, it } from 'vitest';
import {
  emptySearchHistory, findMatches, matchLabel, rememberSearch, replaceMatches, searchHistoryLimit, searchMatchesText,
  searchResultRows, searchText, type SearchOptions,
} from '../../src/application/search-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { findCell, newWorkbook, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';

function ids() {
  let n = 0;
  return () => 'id-' + ++n;
}

function cell(row: number, column: number, content: CellDocument['content'],
  style?: CellDocument['style'], cache?: CellDocument['cache']): CellDocument {
  return { id: 'c-' + row + '-' + column, row, column, content, ...(style && { style }), ...(cache && { cache }) };
}

function book(cells: readonly CellDocument[], extra: readonly CellDocument[] = []) {
  const makeId = ids();
  const document = newWorkbook(makeId);
  const first = document.sheets[0]!;
  const second = { ...first, id: 'sheet-2', name: 'Data' };
  const session = new WorkbookSession({
    ...document,
    sheets: [{ ...first, cells }, { ...second, cells: extra }],
  }, { makeId });
  return { session, first: first.id, second: second.id };
}

const options = (patch: Partial<SearchOptions> = {}): SearchOptions => ({
  query: '10', lookIn: 'values', within: 'sheet', order: 'rows', matchCase: false, wholeCell: false, ...patch,
});

describe('find and replace search', () => {
  it('searches displayed values or stored formulas with ordinal case and whole-cell options', () => {
    const { session, first } = book([
      cell(1, 1, { kind: 'number', value: 10 }, { numberFormat: '0.00' }),
      cell(2, 1, { kind: 'text', value: 'Apple 10' }),
      cell(3, 1, { kind: 'formula', source: '=B1*10' }, undefined, { kind: 'number', value: 20 }),
      cell(4, 1, { kind: 'text', value: 'apple' }),
    ]);
    const document = session.active;
    expect(searchText(findCell(document.sheets[0]!, { row: 1, column: 1 }), 'values')).toBe('10.00');
    expect(searchText(findCell(document.sheets[0]!, { row: 3, column: 1 }), 'formulas')).toBe('=B1*10');
    expect(findMatches(document, [first], options()).map(match => match.row)).toEqual([1, 2]);
    expect(findMatches(document, [first], options({ query: '10.00', wholeCell: true })).map(match => match.row)).toEqual([1]);
    expect(findMatches(document, [first], options({ query: 'APPLE' })).map(match => match.row)).toEqual([2, 4]);
    expect(findMatches(document, [first], options({ query: 'APPLE', matchCase: true }))).toEqual([]);
    expect(findMatches(document, [first], options({ query: 'apple', matchCase: true })).map(match => match.row)).toEqual([4]);
    expect(findMatches(document, [first], options({ query: '=B1*10', lookIn: 'formulas' })).map(match => match.row)).toEqual([3]);
    expect(findMatches(document, [first], options({ query: '20', lookIn: 'formulas' }))).toEqual([]);
    // Find's Formulas option searches formula cells only; Values searches the display of every cell.
    expect(findMatches(document, [first], options({ query: '10', lookIn: 'formulas' })).map(match => match.row)).toEqual([3]);
    expect(findMatches(document, [first], options({ query: '10', lookIn: 'formulas', includeConstants: true })).map(match => match.row))
      .toEqual([1, 2, 3]);
    expect(findMatches(document, [first], options({ query: '=B1*10', lookIn: 'values' }))).toEqual([]);
    expect(findMatches(document, [first], options({ query: '10.00', lookIn: 'formulas' }))).toEqual([]);
    expect(findMatches(document, [first], options({ query: '20', lookIn: 'values' })).map(match => match.row)).toEqual([3]);
    expect(() => findMatches(document, [first], options({ query: '' }))).toThrow('Enter text to find.');
    expect(searchMatchesText('Apple', options({ query: 'app' }))).toBe(true);
  });

  it('orders matches by row or by column inside each sheet and scopes to one or all worksheets', () => {
    const { session, first, second } = book([
      cell(2, 2, { kind: 'text', value: 'x1' }), cell(1, 3, { kind: 'text', value: 'x2' }),
      cell(3, 1, { kind: 'text', value: 'x3' }),
    ], [cell(1, 1, { kind: 'text', value: 'x4' })]);
    const document = session.active;
    const byRows = findMatches(document, [first], options({ query: 'x' }));
    expect(byRows.map(match => [match.row, match.column])).toEqual([[1, 3], [2, 2], [3, 1]]);
    const byColumns = findMatches(document, [first], options({ query: 'x', order: 'columns' }));
    expect(byColumns.map(match => [match.row, match.column])).toEqual([[3, 1], [2, 2], [1, 3]]);
    const both = findMatches(document, [first, second], options({ query: 'x', within: 'workbook' }));
    expect(both.map(match => match.sheetId)).toEqual([first, first, first, second]);
    expect(matchLabel(document, both[0]!, 'sheet')).toBe('C1');
    expect(matchLabel(document, both[3]!, 'workbook')).toBe('Data!A1');
  });

  it('builds the Book, Sheet, Name, Cell, Value and Formula result rows', () => {
    const { session, first } = book([
      cell(2, 3, { kind: 'formula', source: '=SUM(A1:A2)' }, { numberFormat: '0.00' }, { kind: 'number', value: 12.5 }),
    ]);
    const basis = session.capture();
    const rows = searchResultRows(basis, 'Book1', findMatches(basis.document, [first], options({ query: 'SUM', lookIn: 'formulas' })), options());
    expect(rows).toEqual([{ book: 'Book1', sheet: 'Sheet1', name: '', cell: '$C$2', value: '12.50',
      formula: '=SUM(A1:A2)', match: { sheetId: first, row: 2, column: 3 } }]);
  });

  it('replaces every target in one undoable transaction and preserves text versus typed cells', () => {
    const { session, first } = book([
      cell(1, 1, { kind: 'number', value: 10 }, { bold: true }),
      cell(2, 1, { kind: 'text', value: '10 apples' }, { numberFormat: '@' }),
      cell(3, 1, { kind: 'formula', source: '=A1+10' }, undefined, { kind: 'number', value: 20 }),
      cell(4, 1, { kind: 'text', value: 'nothing' }),
    ]);
    const targets = findMatches(session.active, [first], options());
    expect(targets.map(match => match.row)).toEqual([1, 2]);
    session.execute(replaceMatches(session.capture(), targets, options(), '12'));
    const sheet = session.active.sheets[0]!;
    expect(findCell(sheet, { row: 1, column: 1 })).toMatchObject({ content: { kind: 'number', value: 12 }, style: { bold: true } });
    expect(findCell(sheet, { row: 2, column: 1 })?.content).toEqual({ kind: 'text', value: '12 apples' });
    expect(findCell(sheet, { row: 3, column: 1 })?.content).toEqual({ kind: 'formula', source: '=A1+10' });
    expect(findCell(sheet, { row: 4, column: 1 })?.content).toEqual({ kind: 'text', value: 'nothing' });
    expect(session.undoCount).toBe(1);
    expect(session.undo()).toBe(true);
    expect(findCell(session.active.sheets[0]!, { row: 1, column: 1 })?.content).toEqual({ kind: 'number', value: 10 });

    // The Replace tab keeps searching stored source, so constants stay replaceable.
    const formulas = findMatches(session.active, [first], options({ query: '10', lookIn: 'formulas', includeConstants: true }));
    expect(formulas.map(match => match.row)).toEqual([1, 2, 3]);
    session.execute(replaceMatches(session.capture(), formulas,
      options({ query: '10', lookIn: 'formulas', includeConstants: true }), '12'));
    expect(findCell(session.active.sheets[0]!, { row: 3, column: 1 })?.content).toEqual({ kind: 'formula', source: '=A1+12' });

    const whole = findMatches(session.active, [first], options({ query: '12', wholeCell: true }));
    session.execute(replaceMatches(session.capture(), whole, options({ query: '12', wholeCell: true }), 'done'));
    expect(findCell(session.active.sheets[0]!, { row: 1, column: 1 })?.content).toEqual({ kind: 'text', value: 'done' });
    expect(findCell(session.active.sheets[0]!, { row: 2, column: 1 })?.content).toEqual({ kind: 'text', value: '12 apples' });
  });

  it('refuses stale or no-longer-matching targets without touching the workbook', () => {
    const { session, first } = book([cell(1, 1, { kind: 'text', value: 'alpha' })]);
    const targets = findMatches(session.active, [first], options({ query: 'alpha' }));
    session.execute({ kind: 'set-cells', label: 'Edit', edits: [{ target: { sheetId: first, row: 1, column: 1 },
      content: { kind: 'text', value: 'gamma' } }] });
    const before = session.active;
    expect(() => replaceMatches(session.capture(), targets, options({ query: 'alpha' }), 'beta'))
      .toThrow('Search result changed. Find again.');
    expect(session.active).toEqual(before);
    expect(() => replaceMatches(session.capture(), [{ sheetId: first, row: 9, column: 9 }], options({ query: 'x' }), 'y'))
      .toThrow('Search result is no longer available.');
  });

  it('keeps 15 most-recent history entries with exact deduplication', () => {
    expect(rememberSearch([], '')).toEqual([]);
    expect(rememberSearch([], 'alpha')).toEqual(['alpha']);
    expect(rememberSearch(['alpha', 'beta'], 'alpha')).toEqual(['alpha', 'beta']);
    expect(rememberSearch(['alpha', 'beta'], '')).toEqual(['alpha', 'beta']);
    expect(rememberSearch(['alpha'], '', true)).toEqual(['', 'alpha']);
    expect(rememberSearch([], 'x'.repeat(32_768))).toEqual([]);
    expect(rememberSearch([], 'bad\u0000value')).toEqual([]);
    let history: readonly string[] = [];
    for (let index = 0; index < searchHistoryLimit + 3; index++) history = rememberSearch(history, 'query-' + index);
    expect(history).toHaveLength(searchHistoryLimit);
    expect(history[0]).toBe('query-' + (searchHistoryLimit + 2));
    expect(emptySearchHistory).toEqual({ find: [], replace: [] });
  });
});
