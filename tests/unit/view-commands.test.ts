import { describe, expect, it } from 'vitest';
import { freezeFromCaret, setFreeze } from '../../src/application/view-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { documentLimits, newWorkbook } from '../../src/core/workbook';

function session() {
  let n = 0;
  const makeId = () => 'id-' + ++n;
  const document = newWorkbook(makeId);
  const first = document.sheets[0]!;
  const second = { ...first, id: 'sheet-2', name: 'Data' };
  return { workbook: new WorkbookSession({ ...document, sheets: [first, second] }, { makeId }), first: first.id, second: second.id };
}

describe('per-sheet frozen panes', () => {
  it('freezes both axes on the selected worksheet in one transaction', () => {
    const { workbook, first, second } = session();
    const before = workbook.active;
    workbook.execute(setFreeze(workbook.capture(), first, { rows: 4, columns: 3 }));
    expect(workbook.active.sheets[0]!.frozenPanes).toEqual({ rows: 4, columns: 3 });
    expect(workbook.active.sheets[1]!.frozenPanes).toEqual({ rows: 0, columns: 0 });
    expect(workbook.undoCount).toBe(1);
    expect(workbook.undo()).toBe(true);
    expect(workbook.active).toEqual(before);
    expect(workbook.redo()).toBe(true);
    expect(workbook.active.sheets.find(sheet => sheet.id === first)?.frozenPanes).toEqual({ rows: 4, columns: 3 });
    expect(workbook.active.sheets.find(sheet => sheet.id === second)?.frozenPanes).toEqual({ rows: 0, columns: 0 });
  });

  it('unfreezes both axes and refuses out-of-bounds boundaries without changing the workbook', () => {
    const { workbook, first } = session();
    workbook.execute(setFreeze(workbook.capture(), first, { rows: 2, columns: 2 }));
    workbook.execute(setFreeze(workbook.capture(), first, { rows: 0, columns: 0 }));
    expect(workbook.active.sheets[0]!.frozenPanes).toEqual({ rows: 0, columns: 0 });
    const before = workbook.active;
    const undo = workbook.undoCount;
    expect(() => workbook.execute(setFreeze(workbook.capture(), first, { rows: -1, columns: 2 })))
      .toThrow('Frozen pane boundaries are invalid.');
    expect(() => workbook.execute(setFreeze(workbook.capture(), first, { rows: documentLimits.rows, columns: 0 })))
      .toThrow('Frozen pane boundaries are invalid.');
    expect(() => workbook.execute(setFreeze(workbook.capture(), first, { rows: 1.5, columns: 0 })))
      .toThrow('Frozen pane boundaries are invalid.');
    expect(workbook.active).toBe(before);
    expect(workbook.undoCount).toBe(undo);
  });

  it('freezes the rows above and columns left of the caret', () => {
    expect(freezeFromCaret({ row: 1, column: 1 })).toEqual({ rows: 0, columns: 0 });
    expect(freezeFromCaret({ row: 5, column: 3 })).toEqual({ rows: 4, columns: 2 });
    const { workbook, first } = session();
    workbook.execute(setFreeze(workbook.capture(), first, freezeFromCaret({ row: 2, column: 1 })));
    expect(workbook.active.sheets[0]!.frozenPanes).toEqual({ rows: 1, columns: 0 });
    workbook.execute(setFreeze(workbook.capture(), first, { rows: 0, columns: 1 }));
    expect(workbook.active.sheets[0]!.frozenPanes).toEqual({ rows: 0, columns: 1 });
  });
});
