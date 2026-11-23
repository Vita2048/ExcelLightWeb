import { expect, test, vi } from 'vitest';
import { runNativeGeneration } from '../../src/calculation/native-generation';
import { emptySheet, type CellDocument, type WorkbookDocument } from '../../src/core/workbook';

function cell(id: string, row: number, column: number, value: string | number): CellDocument {
  return { id, row, column, content: typeof value === 'number' ? { kind: 'number', value } : { kind: 'formula', source: value } };
}

function book(...cells: CellDocument[]): WorkbookDocument {
  return { id: 'book', formatVersion: 1, dateSystem: '1900', names: [], sheets: [{ ...emptySheet('s', 'Inputs'), cells }] };
}

type Result = ReturnType<typeof runNativeGeneration>;
function apply(document: WorkbookDocument, result: Result): WorkbookDocument {
  return { ...document, sheets: document.sheets.map(sheet => ({
    ...sheet, cells: sheet.cells.map(cell => {
      const update = result.updates.find(({ target }) => target.sheetId === sheet.id &&
        target.row === cell.row && target.column === cell.column);
      return update?.cache ? { ...cell, cache: update.cache } : cell;
    }),
  })) };
}

function values(result: Result) {
  return result.updates.map(({ target, cache }) => [target.sheetId, target.row, target.column, cache?.value]);
}

test('literal edits reuse the dependency graph and evaluate only the range dependents and their downstream closure', () => {
  const document = book(cell('a1', 1, 1, 3), cell('sum', 1, 2, '=SUM(A1:A3)'),
    cell('dependent', 1, 3, '=B1+1'), cell('unrelated', 1, 4, '=2^8'));
  const first = runNativeGeneration(document, 'automatic');
  const cached = apply(document, first);
  const sheet = cached.sheets[0]!;
  const edited = { ...cached, sheets: [{ ...sheet, cells: [...sheet.cells, cell('a3', 3, 1, 10)] }] };
  const second = runNativeGeneration(edited, 'automatic', first.generation);
  expect(second.generation.graph).toBe(first.generation.graph);
  expect(values(second)).toEqual([['s', 1, 2, 13], ['s', 1, 3, 14]]);
  const clock = vi.fn(() => new Date(2026, 10, 18));
  expect(runNativeGeneration(apply(edited, second), 'automatic', second.generation, { clock }).updates).toEqual([]);
  expect(clock).not.toHaveBeenCalled();
});

test('reports the formula cells selected for calculation before evaluating them', () => {
  const document = book(cell('input', 1, 1, 2), cell('first', 1, 2, '=A1+3'),
    cell('second', 1, 3, '=B1+4'), cell('unrelated', 1, 4, '=9'));
  const first = runNativeGeneration(document, 'automatic');
  const cached = apply(document, first);
  const edited = { ...cached, sheets: [{ ...cached.sheets[0]!, cells: cached.sheets[0]!.cells.map(item =>
    item.id === 'input' ? cell('input', 1, 1, 10) : item) }] };
  const progress: string[][] = [];
  const result = runNativeGeneration(edited, 'automatic', first.generation, {}, ids => {
    progress.push([...ids]);
  });
  expect(progress).toEqual([['first', 'second']]);
  expect(values(result)).toEqual([['s', 1, 2, 13], ['s', 1, 3, 17]]);
});

test('SUMIF expanded sum ranges and SUBTOTAL row visibility invalidate dependent nodes', () => {
  const document = book(cell('a1', 1, 1, 1), cell('a2', 2, 1, 1), cell('b1', 1, 2, 5), cell('b2', 2, 2, 6),
    cell('sumif', 1, 3, '=SUMIF(A1:A2,1,B1)'), cell('subtotal', 1, 4, '=SUBTOTAL(109,B1:B2)'),
    cell('unrelated', 1, 5, '=1'));
  const first = runNativeGeneration(document, 'workbook');
  const cached = apply(document, first);
  const sheet = cached.sheets[0]!;
  const edited = { ...cached, sheets: [{ ...sheet, rows: [{ index: 2, hidden: true }],
    cells: sheet.cells.map(item => item.id === 'b2' ? cell('b2', 2, 2, 9) : item) }] };
  const second = runNativeGeneration(edited, 'automatic', first.generation);
  expect(values(second)).toEqual([['s', 1, 3, 14], ['s', 1, 4, 5]]);
  const beforeHide = apply(edited, second);
  const unhidden = { ...beforeHide, sheets: [{ ...beforeHide.sheets[0]!, rows: [] }] };
  expect(values(runNativeGeneration(unhidden, 'automatic', second.generation))).toEqual([['s', 1, 4, 14]]);
});

test('adding and removing a cycle invalidates its members through inactive branches, while IFERROR still handles downstream errors', () => {
  const document = book(cell('a', 1, 1, '=IF(FALSE,B1,7)'), cell('b', 1, 2, 2),
    cell('c', 1, 3, '=IFERROR(A1,99)'), cell('d', 1, 4, '=1'));
  const first = runNativeGeneration(document, 'automatic');
  const edited = apply(document, first);
  const sheet = edited.sheets[0]!;
  const cyclic = { ...edited, sheets: [{ ...sheet, cells: sheet.cells.map(item =>
    item.id === 'b' ? cell('b', 1, 2, '=A1') : item) }] };
  const second = runNativeGeneration(cyclic, 'automatic', first.generation);
  expect(second.generation.graph).not.toBe(first.generation.graph);
  expect(values(second)).toEqual([['s', 1, 2, '#REF!'], ['s', 1, 1, '#REF!'], ['s', 1, 3, 99]]);
  const recovered = apply(cyclic, second);
  const fixed = { ...recovered, sheets: [{ ...recovered.sheets[0]!, cells: recovered.sheets[0]!.cells.map(item =>
    item.id === 'b' ? cell('b', 1, 2, 2) : item) }] };
  expect(values(runNativeGeneration(fixed, 'automatic', second.generation))).toEqual([['s', 1, 1, 7], ['s', 1, 3, 7]]);
});

test('manual observation retains caches and Calculate Sheet evaluates its full precedent closure', () => {
  const inputs = book(cell('input', 1, 1, 2), cell('precedent', 1, 2, '=A1+3'));
  const document = { ...inputs, sheets: [...inputs.sheets,
    { ...emptySheet('current', 'Current'), cells: [cell('current-value', 1, 1, '=Inputs!B1+5')] },
    { ...emptySheet('downstream', 'Downstream'), cells: [cell('downstream-value', 1, 1, '=Current!A1+7')] },
  ] };
  const first = runNativeGeneration(document, 'workbook');
  const cached = apply(document, first);
  const edited = { ...cached, sheets: [{ ...cached.sheets[0]!, cells: [
    cell('input', 1, 1, 10), cached.sheets[0]!.cells[1]!,
  ] }, ...cached.sheets.slice(1)] };
  const clock = vi.fn(() => new Date(2026, 10, 18));
  const observed = runNativeGeneration(edited, 'observe', first.generation, { clock });
  expect(observed.updates).toEqual([]);
  expect(new Set(observed.staleCellIds)).toEqual(new Set(['precedent', 'current-value', 'downstream-value']));
  expect(clock).not.toHaveBeenCalled();
  const scoped = runNativeGeneration(edited, { sheetId: 'current' }, observed.generation, { clock });
  expect(values(scoped)).toEqual([['s', 1, 2, 13], ['current', 1, 1, 18]]);
  expect(scoped.staleCellIds).toEqual(['downstream-value']);
  const cachedScoped = apply(edited, scoped);
  expect(cachedScoped.sheets[2]!.cells[0]!.cache?.value).toBe(17);
  expect(values(runNativeGeneration(cachedScoped, 'automatic', scoped.generation))).toEqual([['downstream', 1, 1, 25]]);
});

test('resolves cross-sheet references when the sheet name contains an escaped apostrophe', () => {
  const source = book(cell('input', 1, 1, 7));
  const document = { ...source, sheets: [
    { ...source.sheets[0]!, name: "O'Brien" },
    { ...emptySheet('main', 'Main'), cells: [cell('result', 1, 1, "='O''Brien'!A1")] },
  ] };
  expect(values(runNativeGeneration(document, 'automatic'))).toEqual([['main', 1, 1, 7]]);
});

test('TODAY uses one injected reading per generation and flows through clean dependents', () => {
  const document = book(cell('today', 1, 1, '=TODAY()'), cell('also-today', 1, 2, '=TODAY()-A1'),
    cell('downstream', 1, 3, '=A1+1'), cell('unrelated', 1, 4, '=1'));
  const clock = vi.fn(() => new Date(2026, 10, 18));
  const first = runNativeGeneration(document, 'automatic', undefined, { clock });
  const day = first.updates[0]!.cache!.value as number;
  expect(clock).toHaveBeenCalledTimes(1);
  clock.mockReturnValue(new Date(2026, 10, 19));
  const second = runNativeGeneration(apply(document, first), 'automatic', first.generation, { clock });
  expect(values(second)).toEqual([['s', 1, 1, day + 1], ['s', 1, 2, 0], ['s', 1, 3, day + 2]]);
  expect(clock).toHaveBeenCalledTimes(2);
});

test('a number-format edit changes active filter keys and invalidates SUBTOTAL without rebuilding the graph', () => {
  const input = cell('a2', 2, 1, 1.25);
  const document = book(input, cell('b2', 2, 2, 8), cell('subtotal', 1, 3, '=SUBTOTAL(9,B2:B2)'));
  const filtered = { ...document, sheets: [{ ...document.sheets[0]!,
    filter: { row: 1, column: 1, rows: 2, columns: 2, criteria: [{ column: 1, allowed: ['1.25'] }] } }] };
  const first = runNativeGeneration(filtered, 'automatic');
  expect(values(first)).toEqual([['s', 1, 3, 8]]);
  const cached = apply(filtered, first);
  const edited = { ...cached, sheets: [{ ...cached.sheets[0]!, cells: cached.sheets[0]!.cells.map(item =>
    item.id === input.id ? { ...item, style: { numberFormat: '0.00%' } } : item) }] };
  const second = runNativeGeneration(edited, 'automatic', first.generation);
  expect(second.generation.graph).toBe(first.generation.graph);
  expect(values(second)).toEqual([['s', 1, 3, 0]]);
});
