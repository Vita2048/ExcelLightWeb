import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../../src/application/workbook-session';
import { runNativeGeneration } from '../../src/calculation/native-generation';
import { calculateNative } from '../../src/calculation/native-calculator';
import { documentLimits, ownDto, type CellDocument, type SheetDocument, type WorkbookDocument } from '../../src/core/workbook';
import { formulaAddress } from '../../src/core/formula-references';
import { validateWorkbook } from '../../src/core/workbook-validation';
import { environment, percentile, record, summarize, type Measurement } from './support';

// Deterministic fixtures at the shared source limits: 2 sheets x 1,000 rows x 100 columns.
const rows = 1_000;
const columns = documentLimits.columns;
const formulaColumns = 5;
const sheets = 2;

function numberCell(id: string, row: number, column: number, value: number): CellDocument {
  return { id, row, column, content: { kind: 'number', value } };
}

function stressSheet(id: string, name: string, idFactory: (part: string) => string): SheetDocument {
  const cells: CellDocument[] = [];
  for (let row = 1; row <= rows; row++) {
    for (let column = 1; column <= columns; column++) {
      const cellId = idFactory(`${row}-${column}`);
      if (column > columns - formulaColumns) {
        const previous = column === columns - formulaColumns + 1 ? `A${row}` : formulaAddress({ row, column: column - 1 });
        cells.push({ id: cellId, row, column, content: { kind: 'formula', source: `=${previous}+${column}` } });
      } else {
        cells.push(numberCell(cellId, row, column, row * 1000 + column));
      }
    }
  }
  return { id, name, cells, merges: [], rows: [], columns: [], frozenPanes: { rows: 0, columns: 0 } };
}

function stressWorkbook(): WorkbookDocument {
  let n = 0;
  const idFactory = (part: string) => `stress-${++n}-${part}`;
  return ownDto({
    id: 'stress-workbook', formatVersion: 1, dateSystem: '1900',
    sheets: [stressSheet('sheet-1', 'Data', idFactory), stressSheet('sheet-2', 'Copy', idFactory)],
    names: [],
  });
}

/** Fixed 1,000-formula dependency chain used by the Development warm-recalculation target. */
function chainWorkbook(): WorkbookDocument {
  const cells: CellDocument[] = [numberCell('chain-1', 1, 1, 1)];
  for (let row = 2; row <= 1_000; row++) {
    cells.push({ id: `chain-${row}`, row, column: 1, content: { kind: 'formula', source: `=A${row - 1}+1` } });
  }
  cells.push({ id: 'chain-sum', row: 1, column: 2, content: { kind: 'formula', source: '=SUM(A1:A1000)' } });
  return ownDto({
    id: 'chain-workbook', formatVersion: 1, dateSystem: '1900',
    sheets: [{ id: 'chain-sheet', name: 'Chain', cells, merges: [], rows: [], columns: [], frozenPanes: { rows: 0, columns: 0 } }],
    names: [],
  });
}

function mediumWorkbook(): WorkbookDocument {
  let n = 0;
  const cells: CellDocument[] = [];
  for (let row = 1; row <= 100; row++) {
    for (let column = 1; column <= documentLimits.columns; column++) {
      cells.push(numberCell(`medium-${++n}`, row, column, row * 1_000 + column));
    }
  }
  return ownDto({
    id: 'medium-workbook', formatVersion: 1, dateSystem: '1900',
    sheets: [{ id: 'medium-sheet', name: 'Medium', cells, merges: [], rows: [], columns: [], frozenPanes: { rows: 0, columns: 0 } }],
    names: [],
  });
}

function applyCaches(document: WorkbookDocument, updates: ReturnType<typeof runNativeGeneration>['updates']): WorkbookDocument {
  const caches = new Map(updates.map(update => [update.target.sheetId + ':' + update.target.row + ':' + update.target.column, update.cache]));
  return ownDto({
    ...document,
    sheets: document.sheets.map(sheet => ({
      ...sheet,
      cells: sheet.cells.map(cell => {
        const cache = caches.get(sheet.id + ':' + cell.row + ':' + cell.column);
        return cache ? { ...cell, cache } : cell;
      }),
    })),
  });
}

function withLiteral(document: WorkbookDocument, sheetIndex: number, row: number, column: number, value: number): WorkbookDocument {
  const sheet = document.sheets[sheetIndex]!;
  return ownDto({
    ...document,
    sheets: document.sheets.map((current, index) => index !== sheetIndex ? current : {
      ...current,
      cells: current.cells.map(cell => cell.row === row && cell.column === column
        ? { ...cell, content: { kind: 'number', value } } : cell),
    }),
  });
}

function time(action: () => void): number {
  const start = performance.now();
  action();
  return performance.now() - start;
}

const samples = (count: number, action: (index: number) => void) => Array.from({ length: count }, (_, index) => time(() => action(index)));

describe('maximum-bound workbook stress', () => {
  it('accepts 200,000 populated cells and 10,000 formulas, then applies single-cell edits the size of worksheet input', () => {
    const workbook = stressWorkbook();
    const cells = workbook.sheets.reduce((total, sheet) => total + sheet.cells.length, 0);
    const formulas = workbook.sheets.reduce((total, sheet) =>
      total + sheet.cells.filter(cell => cell.content.kind === 'formula').length, 0);
    expect(cells).toBe(documentLimits.cells);
    expect(formulas).toBe(documentLimits.formulas);

    let session!: WorkbookSession;
    const construction = time(() => {
      session = new WorkbookSession(workbook, { makeId: (() => { let id = 0; return () => `session-${++id}`; })() });
    });
    const editSamples = samples(20, index => {
      session.execute({ kind: 'set-cells', label: 'Edit cell', edits: [{
        target: { sheetId: session.active.sheets[0]!.id, row: 1, column: 1 },
        content: { kind: 'number', value: index + 2 },
      }] });
    });
    const after = session.active;
    expect(after.sheets.reduce((total, sheet) => total + sheet.cells.length, 0)).toBe(documentLimits.cells);
    expect(after.sheets[0]!.cells.filter(cell => cell.content.kind === 'formula')).toHaveLength(rows * formulaColumns);
    expect(session.undoCount).toBe(20);
    const undo = time(() => { session.undo(); });

    const measurements: Measurement[] = [
      { name: 'validate+construct 200k-cell session (ms)', samples: [construction] },
      { name: 'single-cell session edit at 200k cells (ms)', samples: editSamples },
      { name: 'undo single-cell edit at 200k cells (ms)', samples: [undo] },
    ];
    record('S25-workbook-stress', { environment: environment(), cells, formulas, measurements: summarize(measurements) });
    expect(percentile(editSamples, 0.95)).toBeLessThan(60_000);
  }, 300_000);

  it('recalculates every formula on the maximum-bound workbook without data loss', () => {
    const workbook = stressWorkbook();
    const first = runNativeGeneration(workbook, 'workbook');
    expect(first.updates).toHaveLength(documentLimits.formulas);
    expect(first.updates.every(update => update.cache?.kind === 'number')).toBe(true);
    const caches = new Map(first.updates.map(update => [update.target.sheetId + ':' + update.target.row + ':' + update.target.column, update.cache]));
    // Row 1: A1 = 1001, then the five formula columns each add their own column number.
    const last = caches.get('sheet-1:1:100');
    expect(last).toEqual({ kind: 'number', value: 1_001 + 96 + 97 + 98 + 99 + 100 });

    const cached = applyCaches(workbook, first.updates);
    const edits = Array.from({ length: 5 }, (_, index) => withLiteral(cached, 0, 1, 1, index + 10));
    const warm = edits.map(document => time(() => {
      const result = runNativeGeneration(document, 'automatic', first.generation);
      expect(result.updates.every(update => update.cache?.kind === 'number')).toBe(true);
      // Each row owns a five-cell formula chain fed by its A value.
      expect(result.updates).toHaveLength(formulaColumns);
    }));
    const full = time(() => calculateNative(workbook));
    const measurements: Measurement[] = [
      { name: 'full 10,000-formula maximum-bound generation (ms)', samples: [time(() => runNativeGeneration(workbook, 'workbook'))] },
      { name: 'dirty-chain recalculation after a literal edit (ms)', samples: warm },
      { name: 'synchronous full calculation (ms)', samples: [full] },
    ];
    record('S25-native-stress', { environment: environment(), formulas: documentLimits.formulas, measurements: summarize(measurements) });
    expect(percentile(warm, 0.95)).toBeLessThan(60_000);
  }, 300_000);

  it('keeps the documented warm recalculation target for the 1,000-formula fixture', () => {
    const workbook = chainWorkbook();
    const first = runNativeGeneration(workbook, 'workbook');
    expect(first.updates).toHaveLength(1_000);
    const edits = Array.from({ length: 20 }, (_, index) => withLiteral(workbook, 0, 1, 1, index + 2));
    const warm: number[] = [];
    let generation = first.generation;
    for (const document of edits) {
      let next: ReturnType<typeof runNativeGeneration> | undefined;
      warm.push(time(() => { next = runNativeGeneration(document, 'workbook', generation); }));
      expect(next!.updates).toHaveLength(1_000);
      generation = next!.generation;
    }
    const measurements: Measurement[] = [{ name: 'warm 1,000-formula chain recalculation (ms)', samples: warm }];
    record('S25-native-warm', { environment: environment(), formulas: 1_000, measurements: summarize(measurements) });
    // Development proposes under 500 ms for this fixture; the assertion allows slower CI hosts.
    expect(percentile(warm, 0.95)).toBeLessThan(500);
  }, 120_000);

  it('keeps the 10,000-cell interaction target inside the documented budget', () => {
    const workbook = mediumWorkbook();
    let session!: WorkbookSession;
    const construction = time(() => {
      session = new WorkbookSession(workbook, { makeId: (() => { let id = 0; return () => `medium-session-${++id}`; })() });
    });
    const edits = samples(20, index => {
      session.execute({ kind: 'set-cells', label: 'Edit cell', edits: [{
        target: { sheetId: session.active.sheets[0]!.id, row: 1, column: 1 },
        content: { kind: 'number', value: index + 2 },
      }] });
    });
    const measurements: Measurement[] = [
      { name: 'validate+construct 10k-cell session (ms)', samples: [construction] },
      { name: 'single-cell session edit at 10k cells (ms)', samples: edits },
    ];
    record('S25-workbook-medium', { environment: environment(), cells: 10_000, measurements: summarize(measurements) });
    // Development proposes p95 under 100 ms for worksheet input on a 10,000-cell fixture.
    expect(percentile(edits, 0.95)).toBeLessThan(100);
  }, 120_000);

  it('validates the maximum-bound document and refuses one cell beyond the limit', () => {
    const workbook = stressWorkbook();
    expect(time(() => validateWorkbook(workbook))).toBeGreaterThanOrEqual(0);
    const extra: WorkbookDocument = ownDto({ ...workbook, sheets: workbook.sheets.map((sheet, index) => index !== 0 ? sheet : {
      ...sheet, cells: [...sheet.cells, numberCell('extra-cell', 1_001, 1, 1)],
    }) });
    expect(() => validateWorkbook(extra)).toThrow(/200,000/u);
  }, 120_000);
});
