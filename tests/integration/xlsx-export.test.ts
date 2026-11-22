import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { pasteText } from '../../src/application/editing-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { newWorkbook, type CellDocument, type SheetDocument, type WorkbookDocument } from '../../src/core/workbook';
import { readProbeParts } from '../../src/persistence/bounded-zip';
import { importXlsx, inspectXlsxExport, writeXlsx } from '../../src/persistence/xlsx';

let sequence = 0;
const makeId = () => 'export-id-' + (++sequence);

function workbookOf(cells: readonly CellDocument[], sheet?: Partial<SheetDocument>,
  names: WorkbookDocument['names'] = []): WorkbookDocument {
  return {
    id: makeId(), formatVersion: 1, dateSystem: '1900',
    sheets: [{
      id: makeId(), name: 'Sheet1', cells, rows: [], columns: [], merges: [],
      frozenPanes: { rows: 0, columns: 0 }, ...sheet,
    }],
    names,
  };
}

async function partsOf(blob: Blob): Promise<Map<string, string>> {
  const parts = await readProbeParts(blob);
  return new Map([...parts].map(([name, bytes]) => [name, new TextDecoder().decode(bytes)]));
}

it('reports stable ids, omitted metadata, dimensions, stale caches and missing caches before writing', async () => {
  const base = workbookOf([
    { id: makeId(), row: 1, column: 1, content: { kind: 'number', value: 5 } },
    {
      id: makeId(), row: 2, column: 1, content: { kind: 'formula', source: '=WEBSERVICE(B1)' },
      cache: { kind: 'text', value: 'external' }, requiresExternalCalculation: true,
    },
    { id: makeId(), row: 3, column: 1, content: { kind: 'formula', source: '=SUM(B1:B2)' } },
  ], {
    rows: [{ index: 1, size: 600 }],
    columns: [{ index: 1, size: 2000 }],
    filter: { row: 1, column: 1, rows: 3, columns: 2, criteria: [] },
  });
  const sheetId = base.sheets[0]!.id;
  const document: WorkbookDocument = {
    ...base,
    names: [
      { id: makeId(), name: 'Total', area: { sheetId, row: 1, column: 1, rows: 1, columns: 1 } },
      {
        id: makeId(), name: 'SalesTable',
        area: { sheetId, row: 1, column: 1, rows: 2, columns: 2 },
        isTable: true, design: { style: 'Blue', headerRow: true, totalRow: false }, headerLabels: ['A', 'B'],
      },
    ],
  };
  const report = inspectXlsxExport(document);
  for (const code of ['stable-ids', 'unevaluated-formulas', 'omitted-objects', 'defined-names',
    'tables', 'unsupported-formulas', 'filter-freeze', 'python-snapshots', 'dimension-limits',
    'column-width', 'cache-freshness', 'missing-cache']) {
    expect(report.has(code), code).toBe(true);
  }
  expect(report.issues.find(issue => issue.code === 'missing-cache')?.detail).toContain('Sheet1!A3');
  expect(report.has('macros')).toBe(false);
});

it('writes formula text and caches without evaluating or inventing zero, and asks Excel to recalculate', async () => {
  const document = workbookOf([
    { id: makeId(), row: 1, column: 1, content: { kind: 'formula', source: '=SUM(B1:B2)' }, cache: { kind: 'number', value: 12 } },
    { id: makeId(), row: 2, column: 1, content: { kind: 'formula', source: '=NOW()' } },
    { id: makeId(), row: 1, column: 2, content: { kind: 'number', value: 5 } },
    { id: makeId(), row: 2, column: 2, content: { kind: 'number', value: 7 } },
  ]);
  const parts = await partsOf((await writeXlsx(document)).blob);
  const sheet = parts.get('xl/worksheets/sheet1.xml')!;
  expect(sheet).toContain('<f>SUM(B1:B2)</f><v>12</v>');
  const uncached = /<c r="A2"[^>]*>(.*?)<\/c>/u.exec(sheet)?.[1];
  expect(uncached).toBe('<f>NOW()</f>');
  expect(uncached).not.toContain('<v>');
  expect(parts.get('xl/workbook.xml')).toContain('calcMode="manual"');
  expect(parts.get('xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
});

it('exports only materialized Python values and never its source, bindings or execution', async () => {
  const source = 'import pandas as pd\nresult = 7';
  const blockId = makeId();
  const uncachedBlockId = makeId();
  const document = workbookOf([
    {
      id: makeId(), row: 1, column: 1,
      content: {
        kind: 'python', blockId, source,
        block: {
          id: blockId, source, bindings: [], status: 'succeeded', hasResult: true,
          lastValue: { kind: 'number', value: 7 }, spillActive: false, placements: [],
        },
      },
      cache: { kind: 'number', value: 7 },
    },
    {
      id: makeId(), row: 2, column: 1,
      content: {
        kind: 'python', blockId: uncachedBlockId, source,
        block: {
          id: uncachedBlockId, source, bindings: [], status: 'never', hasResult: false,
          lastValue: null, spillActive: false, placements: [],
        },
      },
    },
  ]);
  const { blob, report } = await writeXlsx(document);
  const parts = await partsOf(blob);
  for (const text of parts.values()) {
    expect(text).not.toContain('pandas');
    expect(text).not.toContain('result = 7');
  }
  expect(parts.get('xl/worksheets/sheet1.xml')).toContain('<v>7</v>');
  expect(report.has('python-snapshots')).toBe(true);
});

it('omits filters and defined names from the output while reporting them', async () => {
  const document = workbookOf([
    { id: makeId(), row: 1, column: 1, content: { kind: 'number', value: 1 } },
  ]);
  const sheetId = document.sheets[0]!.id;
  const withMetadata: WorkbookDocument = {
    ...document,
    names: [{ id: makeId(), name: 'Total', area: { sheetId, row: 1, column: 1, rows: 1, columns: 1 } }],
    sheets: [{ ...document.sheets[0]!, filter: { row: 1, column: 1, rows: 1, columns: 1, criteria: [] } }],
  };
  const parts = await partsOf((await writeXlsx(withMetadata)).blob);
  expect(parts.get('xl/workbook.xml')).not.toContain('definedNames');
  expect(parts.get('xl/worksheets/sheet1.xml')).not.toContain('autoFilter');
});

it('round trips the showcase workbook content, caches, styles, dimensions and freeze', async () => {
  const source = new Blob([await readFile(new URL('../fixtures/sales/showcase_data.xlsx', import.meta.url))]);
  const first = await importXlsx(source);
  const second = await importXlsx((await writeXlsx(first.document)).blob);
  expect(second.document.sheets.map(sheet => sheet.name)).toEqual(first.document.sheets.map(sheet => sheet.name));
  for (const [index, sheet] of first.document.sheets.entries()) {
    const other = second.document.sheets[index]!;
    expect(other.cells.map(cell => [cell.row, cell.column, cell.content, cell.cache, cell.style]))
      .toEqual(sheet.cells.map(cell => [cell.row, cell.column, cell.content, cell.cache, cell.style]));
    expect(other.rows).toEqual(sheet.rows);
    expect(other.columns).toEqual(sheet.columns);
    expect(other.frozenPanes).toEqual(sheet.frozenPanes);
    expect(other.merges).toEqual(sheet.merges);
  }
});

it('exporting does not mark the workbook saved or clear its dirty state', async () => {
  const session = new WorkbookSession(newWorkbook(makeId), { makeId });
  const basis = session.capture();
  session.execute(pasteText(basis, { sheetId: session.active.sheets[0]!.id, row: 1, column: 1, rows: 1, columns: 1 }, 'Draft', 'all', 'en-US', makeId));
  expect(session.isDirty).toBe(true);
  const savedBefore = session.savedIdentity;
  const undoBefore = session.undoCount;
  const revisionBefore = session.revision;

  await writeXlsx(session.active);

  expect(session.isDirty).toBe(true);
  expect(session.savedIdentity).toBe(savedBefore);
  expect(session.undoCount).toBe(undoBefore);
  expect(session.revision).toBe(revisionBefore);
  expect(session.active.sheets[0]!.cells[0]!.content).toEqual({ kind: 'text', value: 'Draft' });
});

it('writes exports for the desktop XLSX oracle when S19_XLSX_OUT is set', async () => {
  const output = process.env.S19_XLSX_OUT;
  if (!output) return;
  await mkdir(output, { recursive: true });
  for (const [name, fixture] of [
    ['showcase-web.xlsx', '../fixtures/sales/showcase_data.xlsx'],
    ['styles-web.xlsx', '../fixtures/interchange/DesktopStyles.xlsx'],
  ] as const) {
    const { document } = await importXlsx(new Blob([await readFile(new URL(fixture, import.meta.url))]));
    const { blob } = await writeXlsx(document);
    await writeFile(join(output, name), new Uint8Array(await blob.arrayBuffer()));
  }
});
